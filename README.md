# m365-license-optimizer
Checklist and read-only PowerShell script to find inactive licensed users and cut Microsoft 365 license costs.
<#
.SYNOPSIS
    Reports licensed Microsoft 365 users with no recent sign-in activity.

.DESCRIPTION
    Read-only. Uses Microsoft Graph to list users who hold at least one license
    and have not signed in within the given number of days (or never signed in).
    Exports the results to CSV for manual review. Does NOT change any license.

.PARAMETER InactiveDays
    Number of days without sign-in before a user is flagged. Default: 90.

.PARAMETER OutputPath
    Path of the CSV report. Default: ./inactive-licensed-users.csv

.EXAMPLE
    ./Get-InactiveLicensedUsers.ps1 -InactiveDays 60 -OutputPath ./report.csv
#>

[CmdletBinding()]
param(
    [int]$InactiveDays = 90,
    [string]$OutputPath = "./inactive-licensed-users.csv"
)

$ErrorActionPreference = "Stop"

# Connect with read-only scopes
Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All", "Organization.Read.All" -NoWelcome

# Map license SKU IDs to readable names
$skuMap = @{}
Get-MgSubscribedSku | ForEach-Object { $skuMap[$_.SkuId] = $_.SkuPartNumber }

$cutoff = (Get-Date).AddDays(-$InactiveDays)

Write-Host "Fetching users (this can take a while in large tenants)..."
$users = Get-MgUser -All `
    -Property Id, DisplayName, UserPrincipalName, AccountEnabled, AssignedLicenses, SignInActivity

$report = foreach ($user in $users) {
    if ($user.AssignedLicenses.Count -eq 0) { continue }

    $lastSignIn = $user.SignInActivity.LastSignInDateTime
    if ($lastSignIn -and $lastSignIn -gt $cutoff) { continue }

    [PSCustomObject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        AccountEnabled    = $user.AccountEnabled
        LastSignIn        = $lastSignIn
        DaysInactive      = if ($lastSignIn) { [int]((Get-Date) - $lastSignIn).TotalDays } else { $null }
        Licenses          = ($user.AssignedLicenses | ForEach-Object {
                                if ($skuMap.ContainsKey($_.SkuId)) { $skuMap[$_.SkuId] } else { $_.SkuId }
                             }) -join "; "
    }
}

$report | Sort-Object DaysInactive -Descending | Export-Csv -Path $OutputPath -NoTypeInformation

Write-Host "Done. $(@($report).Count) inactive licensed user(s) found. Report saved to $OutputPath"
Write-Host "Review the list with HR or team leads before reclaiming any licenses."
