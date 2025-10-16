package main

import (
	"bytes"
	"crypto/rand"
	"encoding/json"
	"fmt"
	"io"
	"math/big"
	"net/http"
	"net/url"
	"os"
	"strings"
)

// Environment variables
var (
	tenantID = os.Getenv("TENANT_ID")
	clientID = os.Getenv("CLIENT_ID")
	driveID  = os.Getenv("DRIVE_ID")
	fileID   = os.Getenv("FILE_ID")
	table    = "OfficeForms.Table"
)

// Mandatory columns
var mandatoryColumns = []string{
	"First Name as appears in government-issued ID",
	"Last Name as appears in government-issued ID",
}

// All columns from sheet
var headers = []string{
	"Id", "Start time", "Completion time", "Email", "Name", "Language",
	"First Name as appears in government-issued ID",
	"Last Name as appears in government-issued ID",
	"\"Preferred Name\n\"",
	"Start Date", "Account Type", "Job Role", "Highest Teaching Grade",
	"Private E-mail Address (will only be used to deliver the temporary access information)",
	"Employment Type", "Class",
	"Parent or Legal Guardian #1 FIRST name (or preferred first name)",
	"Parent or Legal Guardian #1 LAST name",
	"Parent or Legal Guardian #1 - E-mail Address",
	"Parent or Legal Guardian #1 - Phone Number",
	"Add another Parent or Legal Guardian?",
	"Parent or Legal Guardian #2 FIRST name (or preferred name)",
	"Parent or Legal Guardian #2 LAST name",
	"Parent or Legal Guardian #2 - E-mail Address",
	"Parent or Legal Guardian #2 - Phone Number",
	"Wi-Fi Account Required?",
}

// getAccessToken fetches a client credentials token
func getAccessToken() (string, error) {
	clientSecret := os.Getenv("AZURE_CLIENT_SECRET")
	if clientSecret == "" {
		return "", fmt.Errorf("AZURE_CLIENT_SECRET not set in environment")
	}

	form := url.Values{}
	form.Set("grant_type", "client_credentials")
	form.Set("client_id", clientID)
	form.Set("client_secret", clientSecret)
	form.Set("scope", "https://graph.microsoft.com/.default")

	tokenURL := "https://login.microsoftonline.com/" + tenantID + "/oauth2/v2.0/token"
	resp, err := http.Post(tokenURL, "application/x-www-form-urlencoded", strings.NewReader(form.Encode()))
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	var result map[string]interface{}
	if err := json.Unmarshal(body, &result); err != nil {
		return "", err
	}

	if token, ok := result["access_token"].(string); ok {
		return token, nil
	}
	return "", fmt.Errorf("no access_token in response: %v", result)
}

// fetchRows reads all rows from the Excel table
func fetchRows(token string) ([]interface{}, error) {
	url := fmt.Sprintf("https://graph.microsoft.com/v1.0/drives/%s/items/%s/workbook/tables/%s/rows", driveID, fileID, table)
	req, _ := http.NewRequest("GET", url, nil)
	req.Header.Set("Authorization", "Bearer "+token)

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	var data map[string]interface{}
	if err := json.Unmarshal(body, &data); err != nil {
		return nil, err
	}

	if rows, ok := data["value"].([]interface{}); ok && len(rows) > 0 {
		fmt.Printf("Found %d rows in %s\n", len(rows), table)
		return rows, nil
	}
	return nil, fmt.Errorf("no rows found in table")
}

// generateRandomPassword creates a secure random password
func generateRandomPassword(length int) string {
	const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*"
	var password strings.Builder
	for i := 0; i < length; i++ {
		n, _ := rand.Int(rand.Reader, big.NewInt(int64(len(chars))))
		password.WriteByte(chars[n.Int64()])
	}
	return password.String()
}

// checkIfProcessed verifies if a row was already processed (stored in processed.log)
func checkIfProcessed(identifier string) bool {
	data, err := os.ReadFile("processed.log")
	if err != nil {
		return false
	}
	return strings.Contains(string(data), identifier)
}

// logProcessedRow writes a processed row identifier to processed.log
func logProcessedRow(identifier string) {
	f, err := os.OpenFile("processed.log", os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err == nil {
		defer f.Close()
		f.WriteString(identifier + "\n")
	}
}

// createUserFromRow creates an Entra user and triggers parent contact + email
func createUserFromRow(token string, row []interface{}) (string, string, error) {
	data := map[string]string{}
	for i, val := range row {
		if i < len(headers) {
			data[headers[i]] = fmt.Sprintf("%v", val)
		}
	}

	for _, col := range mandatoryColumns {
		if strings.TrimSpace(data[col]) == "" {
			fmt.Println("Skipping row due to missing mandatory field:", col)
			return "", "", nil
		}
	}

	firstName := data["First Name as appears in government-issued ID"]
	lastName := data["Last Name as appears in government-issued ID"]
	preferredName := data["\"Preferred Name\n\""]
	accountType := data["Account Type"]
	class := data["Class"]

	displayName := preferredName
	if displayName == "" {
		displayName = firstName + " " + lastName
	}

	username := strings.ToLower(firstName + "." + lastName)
	upn := username + "@ldv-muenchen.de"
	password := generateRandomPassword(16)

	// Skip if user was already processed
	if checkIfProcessed(upn) {
		fmt.Println("Skipping already processed user:", upn)
		return upn, password, nil
	}

	user := map[string]interface{}{
		"accountEnabled":    true,
		"displayName":       displayName,
		"givenName":         firstName,
		"surname":           lastName,
		"mailNickname":      username,
		"userPrincipalName": upn,
		"jobTitle":          accountType,
		"department":        class,
		"passwordProfile": map[string]interface{}{
			"forceChangePasswordNextSignIn": true,
			"password":                      password,
		},
	}

	userJSON, _ := json.Marshal(user)
	req, _ := http.NewRequest("POST", "https://graph.microsoft.com/v1.0/users", bytes.NewBuffer(userJSON))
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return "", "", err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	fmt.Println("Graph Response:", string(body))

	if resp.StatusCode >= 300 {
		return "", "", fmt.Errorf("failed to create user: %s", resp.Status)
	}

	fmt.Printf("User created: %s\n", upn)
	logProcessedRow(upn)

	// after user creation: handle parent contacts
	if err := HandleParents(token, row, upn, password, "it-admin@ldv-muenchen.de"); err != nil {
		fmt.Printf("Error handling parents for %s: %v\n", upn, err)
	}

	return upn, password, nil
}

func main() {
	token, err := getAccessToken()
	if err != nil {
		panic(err)
	}

	rows, err := fetchRows(token)
	if err != nil {
		panic(err)
	}

	for _, r := range rows {
		row := r.(map[string]any)["values"].([]any)[0].([]any)
		upn, _, err := createUserFromRow(token, row)
		if err != nil {
			fmt.Println("Error creating user:", err)
		} else if upn != "" {
			fmt.Printf("Processed user: %s\n", upn)
		}
	}
}
