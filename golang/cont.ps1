param(
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][string]$Email,
    [Parameter(Mandatory = $true)][string]$Class
)

# Service account UPN from environment variable
$ServiceUser = $env:SERVICE_ACCOUNT_UPN
if (-not $ServiceUser) {
    Write-Error "Missing environment variable: SERVICE_ACCOUNT_UPN"
    exit 1
}

# Retrieve service account password securely from SecretManagement
try {
    $SecurePassword = Get-Secret -Name "ExchangeSvcPassword"
    if (-not $SecurePassword) {
        Write-Error "Could not retrieve secret 'ExchangeSvcPassword'."
        exit 1
    }
} catch {
    Write-Error "Failed to retrieve password from SecretManagement: $_"
    exit 1
}

# Convert to PSCredential
$Cred = New-Object System.Management.Automation.PSCredential($ServiceUser, $SecurePassword)

# Import Exchange Online module
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
} catch {
    Write-Error "Failed to import ExchangeOnlineManagement module: $_"
    exit 1
}

# Connect using service account credentials (non-interactive)
try {
    Connect-ExchangeOnline -Credential $Cred -ShowBanner:$false
} catch {
    Write-Error "Failed to connect to Exchange Online: $_"
    exit 1
}

# Check if mail contact exists
try {
    $existing = Get-MailContact -Filter "ExternalEmailAddress -eq '$Email'" -ErrorAction SilentlyContinue
} catch {
    Write-Error "Failed to check existing contacts: $_"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# Create mail contact if it doesn't exist
if (-not $existing) {
    Write-Host "Creating new mail contact for $Name $Email"
    try {
        New-MailContact -Name $Name -ExternalEmailAddress $Email -ErrorAction Stop
        Write-Host "Mail contact created successfully."
    } catch {
        Write-Error "Failed to create mail contact: $_"
        Disconnect-ExchangeOnline -Confirm:$false
        exit 1
    }
} else {
    Write-Host "Mail contact already exists for $Email"
}

# Map class to distribution group
$classToDG = @{
    "GS-1A"  = "pg_01a@ldv-muenchen.de"
    "GS-1B"  = "pg_01a@ldv-muenchen.de"
    "GS-2A"  = "pg_02a@ldv-muenchen.de"
    "GS-2B"  = "pg_02b@ldv-muenchen.de"
    "GS-3A"  = "pg_03a@ldv-muenchen.de"
    "GS-3B"  = "pg_03b@ldv-muenchen.de"
    "GS-4A"  = "pg_04a@ldv-muenchen.de"
    "GS-4B"  = "pg_04b@ldv-muenchen.de"
    "GYM-5A" = "pg_05a@ldv-muenchen.de"
    "GYM-5B" = "pg_05b@ldv-muenchen.de"
    "GYM-6A" = "pg_06a@ldv-muenchen.de"
    "GYM-6B" = "pg_06b@ldv-muenchen.de"
    "GYM-7A" = "pg_07a@ldv-muenchen.de"
    "GYM-7B" = "pg_07b@ldv-muenchen.de"
    "GYM-8"  = "pg_08@ldv-muenchen.de"
    "GYM-9"  = "pg_09@ldv-muenchen.de"
    "GYM-10" = "pg_10@ldv-muenchen.de"
    "GYM-11" = "pg_11@ldv-muenchen.de"
    "GYM-12" = "pg_12@ldv-muenchen.de"
    "GYM-13" = "pg_13@ldv-muenchen.de"
}

$groupEmail = $classToDG[$Class]
if (-not $groupEmail) { $groupEmail = "default@ldv-muenchen.de" }

# Add contact to distribution group
try {
    Add-DistributionGroupMember -Identity $groupEmail -Member $Email -ErrorAction Stop
    Write-Host "Added $Email to distribution group $groupEmail."
} catch {
    Write-Error "Failed to add $Email to distribution group $($groupEmail): $_"
}

# Disconnect session (non-interactive)
try {
    Disconnect-ExchangeOnline -Confirm:$false
} catch {
    Write-Error "Failed to disconnect Exchange session: $_"
}
