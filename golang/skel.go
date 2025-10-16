package main

import (
	"fmt"
	"os"
	"os/exec"
)

// HandleParents processes parent info from a form submission
func HandleParents(token string, row []interface{}, studentUPN, studentPassword, senderEmail string) error {
	data := map[string]string{}
	for i, val := range row {
		if i < len(headers) {
			data[headers[i]] = fmt.Sprintf("%v", val)
		}
	}

	class := data["Class"]

	parents := []struct {
		Name  string
		Email string
	}{
		{data["Parent or Legal Guardian #1 FIRST name (or preferred first name)"] + " " + data["Parent or Legal Guardian #1 LAST name"], data["Parent or Legal Guardian #1 - E-mail Address"]},
		{data["Parent or Legal Guardian #2 FIRST name (or preferred name)"] + " " + data["Parent or Legal Guardian #2 LAST name"], data["Parent or Legal Guardian #2 - E-mail Address"]},
	}

	for _, parent := range parents {
		if parent.Email != "" {
			if err := createMailContact(parent.Name, parent.Email, class); err != nil {
				return fmt.Errorf("failed to create parent contact: %v", err)
			}
			if err := SendParentEmail(token, parent.Email, studentUPN, studentPassword); err != nil {
				return fmt.Errorf("failed to send email: %v", err)
			}
		}
	}

	return nil
}

// createMailContact calls the PowerShell script to create the Exchange contact and add to DG
func createMailContact(name, email, class string) error {
	if name == "" || email == "" {
		return fmt.Errorf("name and email are required")
	}

	scriptPath := `C:\Users\Novus\Documents\golang\cont.ps1`

	cmd := exec.Command("pwsh", "-File", scriptPath, "-Name", name, "-Email", email, "-Class", class)

	cmd.Env = append(os.Environ(),
		"SERVICE_ACCOUNT_UPN="+os.Getenv("SERVICE_ACCOUNT_UPN"),
	)

	out, err := cmd.CombinedOutput()
	if err != nil {
		return fmt.Errorf("PowerShell script failed: %v\nOutput: %s", err, string(out))
	}

	fmt.Printf("PowerShell output:\n%s\n", string(out))
	return nil
}
