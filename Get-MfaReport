<#
SYNOPSIS
Creates a simple report showing MFA and SSPR registration status for all users in your Microsoft 365 tenant.

WHAT THIS REPORT TELLS YOU
It helps you understand:
- Who has MFA turned on
- Who has Self-Service Password Reset set up
- How users are signing in and what methods they use

WHAT YOU NEED BEFORE RUNNING IT
- Microsoft Graph PowerShell installed
- Permission to read user and report data:
  Reports.Read.All and User.Read.All
#>

# ---------------------------
# Step 1: Sign in to Microsoft Graph
# ---------------------------
Connect-MgGraph -Scopes "Reports.Read.All","User.Read.All"

# ---------------------------
# Step 2: Get data from Microsoft Graph
# ---------------------------
$uri = "https://graph.microsoft.com/beta/reports/authenticationMethods/userRegistrationDetails"

$allUsers = @()

# Keep pulling data until everything is retrieved
while ($uri) {
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri
    $allUsers += $response.value
    $uri = $response.'@odata.nextLink'
}

# ---------------------------
# Step 3: Clean and organize the data
# ---------------------------
$report = foreach ($u in $allUsers) {
    [PSCustomObject]@{
        UserPrincipalName = $u['userPrincipalName']
        DisplayName       = $u['userDisplayName']

        MFARegistered     = $u['isMfaRegistered']
        SSPRRegistered    = $u['isSsprRegistered']
        SSPREnabled       = $u['isSsprEnabled']

        AuthenticationMethods = ($u['methodsRegistered'] -join ', ')
        PreferredSignInMethod = $u['userPreferredMethodForSecondaryAuthentication']

        IsAdmin           = $u['isAdmin']
        UserType          = $u['userType']
        CanUseMFA         = $u['isMfaCapable']
    }
}

# ---------------------------
# Step 4: Save report to a file
# ---------------------------
$outputPath = ".\MFA_SSPR_User_Report.csv"

$report |
    Sort-Object UserPrincipalName |
    Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Host "Report is ready! You can find it here: $outputPath"
