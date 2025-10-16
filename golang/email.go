package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"os"
	"strings"
)

const schoolLogoURL = "https://i.imgur.com/0QV6BIW.png"

// SendParentEmail sends the student credentials email to a parent
func SendParentEmail(token, parentEmail, studentUPN, studentPassword string) error {
	serviceAccountUPN := os.Getenv("SERVICE_ACCOUNT_UPN")
	if serviceAccountUPN == "" {
		return fmt.Errorf("SERVICE_ACCOUNT_UPN environment variable not set")
	}

	emailTemplate := `
<html>
<head>
<meta charset="UTF-8">
<style>
    body { font-family: Segoe UI, Arial, sans-serif; font-size: 14px; color: #333; }
    .section { margin-bottom: 20px; }
    .lang-title { font-weight: bold; font-size: 16px; margin-bottom: 5px; }
    .credentials { background-color: #f2f2f2; padding: 10px; border-radius: 5px; }
    .credentials p { margin: 5px 0; font-weight: bold; }
    a { color: #2a72de; text-decoration: none; }
</style>
</head>
<body>

<div style="text-align: center; padding: 20px 0;">
    <img src="$schoolLogoUrl" alt="School Logo" style="max-height: 100px;">
</div>

<div class="section">
    <div class="lang-title">Benvenuto/a!</div>
    <p>La nostra scuola utilizza <b>Microsoft Teams</b>, <b>Microsoft Outlook</b> e <b>Office 365</b> come piattaforma principale di apprendimento e organizzazione.</p>
    <p>Offriamo agli studenti, al personale scolastico e alle famiglie l'opportunità di installare gratuitamente Office su un massimo di 4 dispositivi.</p>
    <p>Le applicazioni Office sono sempre accessibili tramite browser.</p>
    <p>Qui di seguito ti mandiamo i tuoi dati di accesso personali. Ti verrà richiesto di cambiare la password.</p>
</div>

<div class="section">
    <div class="lang-title">Willkommen!</div>
    <p>Unsere Schule verwendet <b>Microsoft Teams</b>, <b>Microsoft Outlook</b> und <b>Office 365</b>.</p>
    <p>Schüler und Familien können Office kostenlos auf bis zu 4 Geräten installieren.</p>
    <p>Office ist auch im Browser nutzbar.</p>
    <p>Deine persönlichen Zugangsdaten findest du unten. Das Passwort muss beim ersten Login geändert werden.</p>
</div>

<div class="section credentials">
    <p><b>Username &amp; E-Mail:</b> $userPrincipalName</p>
    <p><b>Temporary Password:</b> $password</p>
</div>

<div class="section">
    <p><b>Primo accesso / Erste Anmeldung:</b> <a href="https://www.office.com">https://www.office.com</a></p>
    <p><b>Teams App:</b> <a href="https://www.microsoft.com/it-it/microsoft-teams/download-app">Download</a></p>
</div>

</body>
</html>
`

	// Replace placeholders
	emailBody := strings.ReplaceAll(emailTemplate, "$schoolLogoUrl", schoolLogoURL)
	emailBody = strings.ReplaceAll(emailBody, "$userPrincipalName", studentUPN)
	emailBody = strings.ReplaceAll(emailBody, "$password", studentPassword)

	// Prepare Graph API email request
	mail := map[string]interface{}{
		"message": map[string]interface{}{
			"subject": "Accesso account studente / Schulerkonto-Zugang",
			"body": map[string]string{
				"contentType": "HTML",
				"content":     emailBody,
			},
			"toRecipients": []map[string]interface{}{
				{"emailAddress": map[string]string{"address": parentEmail}},
			},
			"from": map[string]interface{}{
				"emailAddress": map[string]string{"address": "it-admin@ldv-muenchen.de"},
			},
			"internetMessageHeaders": []map[string]string{
				{"name": "X-Encrypt", "value": "true"},
			},
		},
		"saveToSentItems": "true",
	}

	jsonMail, _ := json.Marshal(mail)
	req, _ := http.NewRequest(
		"POST",
		fmt.Sprintf("https://graph.microsoft.com/v1.0/users/%s/sendMail", serviceAccountUPN),
		bytes.NewBuffer(jsonMail),
	)
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")

	client := &http.Client{}
	resp, err := client.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	body, _ := io.ReadAll(resp.Body)
	fmt.Println("Email Sent Successfully.", string(body))

	if resp.StatusCode >= 300 {
		return fmt.Errorf("graph API error sending email: %v", resp.Status)
	}
	return nil
}
