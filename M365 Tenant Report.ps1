#References
# https://learn.microsoft.com/en-us/graph/overview?context=graph%2Fapi%2F1.0&view=graph-rest-1.0
# https://practical365.com/mfa-status-user-accounts/ 
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-UserPasswordChanges.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/Report-ExpiringAppSecrets.PS1 
# https://gist.github.com/svarukala/1f473184e5ec8b36d3444ab8f5ce85e8/revisions 
# https://gcits.com/knowledge-base/export-customers-microsoft-secure-scores-to-csv-and-html-reports/ 
# https://o365info.com/export-conditional-access-policy/ 
# https://medium.com/@mozzeph/translate-microsoft-365-license-guids-to-product-names-in-powershell-e8fa373ace16 
# https://practical365.com/entra-id-multifactor-authentication-reaches-38-percent/

# Check if the Microsoft Graph module is installed
function Install-ModuleIfNotExists {
    param (
        [Parameter(Mandatory=$true)]
        [string] $ModuleName
    )
    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName is not installed. Installing now..."
        Install-Module -Name $ModuleName -Scope CurrentUser -Force
    }
}

# Install required modules
Install-ModuleIfNotExists -ModuleName "Microsoft.Graph"
Install-ModuleIfNotExists -ModuleName "Microsoft.Graph.Beta"
Install-ModuleIfNotExists -ModuleName "PsWriteHTML"

# Import modules
Import-Module -Name "PsWriteHTML"
Connect-MgGraph -Scopes AuditLog.Read.All, Directory.Read.All, UserAuthenticationMethod.Read.All, Policy.Read.All, Application.Read.All,SecurityEvents.Read.All, Reports.Read.All, Organization.Read.All -nowelcome
#Report Variables

$MSPLogo    = "Your Logo URL"
$FavIcon    = "Ýour Favicon URL"
$OrgName  = (Get-MgOrganization).DisplayName

$ExchangeUsageCSVPath           = './ExchangeUsage.csv'
$OneDriveUsageCSVPath           = './OneDriveUsage.csv'
$SharePointUsageCSVPath         = './SharePointUsage.csv'
$UserDetailCSVPath              = './UserDetail.csv'
$MailboxDetailCSVPath           = './MailboxDetail.csv'
$OneDriveDetailCSVPath          = './OneDriveDetail.csv'
$SharePointDetailCSVPath        = './SharePointDetail.csv'
$O365GroupsActivityCSVPath      = './O365GroupsActivity.csv'
$SharePointActivityCSVPath      = './SharepointActivity.csv'
$OneDriveActivityCSVPath        = './OneDriveActivity.csv'
$TeamsActivityCSVPath           = './TeamsActivity.csv'
$TeamsUserActivityCSVPath       = './TeamsActivityDist.csv'
$O365ActivationUserCSVPath      = './O365ActivationUser.csv'
$M365AppUserDetailCSVPath       = './M365AppUserDetail.csv'

$ReportFile                     = "./$OrgName.html"



$Headers = @{ConsistencyLevel="Eventual"}  
$TenantId = (Get-MgOrganization).Id
$StartDate = (Get-Date).AddDays(-30)
$StartDateS = (Get-Date $StartDate -Format s) + "Z"

Write-Host "Looking for sign-in records..."
try {
[array]$AuditRecords = Get-MgBetaAuditLogSignIn -Top 10000 `
  -Filter "(CreatedDateTime ge $StartDateS) and (signInEventTypes/any(t:t eq 'interactiveuser')) and (usertype eq 'Member')" -ErrorAction Stop
If (!$AuditRecords) {
    Write-Host "No sign-in records found - exiting"
    
}

# Eliminate any member sign-ins from other tenants
$AuditRecords = $AuditRecords | Where-Object HomeTenantId -match $TenantId

 # Get MFA data
 [array]$MFAData = Get-MgReportAuthenticationMethodUserRegistrationDetail -top 20000
 $MFAData = $MFAData | Where-Object {$_.userType -eq 'Member'} -ErrorAction Stop

Write-Host "Finding user accounts to check..."
[array]$Users = Get-MgBetaUser -All -Sort 'displayName' `
    -Filter "userType eq 'Member'" -consistencyLevel eventual -CountVariable UsersFound `
    -Property Id, displayName, signInActivity, userPrincipalName, AccountEnabled, AssignedLicenses -ErrorAction Stop

Write-Host ("Checking {0} sign-in audit records for {1} user accounts..." -f $AuditRecords.count, $Users.count)
[int]$MFAUsers = 0

} catch {
    $catchBlockExecuted = $true
    Write-Host "An error occurred: $_"
    Write-Host "Falling back to MSOLOnline module..."
   
    if (-not (Get-Module -ListAvailable -Name MSOnline)) {
        Write-Host "MSOLOnline module not found. Installing..."
        Install-Module -Name MSOnline -Force
    }

        Import-Module MSOnline
        Write-Host "Sign in to your account" -ForegroundColor Yellow
        Connect-MsolService
        $Users = Get-MgBetaUser -All | Where-Object { $_.UserType -eq 'Member' }
        $Report = New-Object System.Collections.Generic.List[Object]
        ForEach ($User in $Users) {
            $UserSignInStatus = if ($User.AccountEnabled -eq $true) {"Allowed"} else {"Blocked"}
            $MsolUser = Get-MsolUser -UserPrincipalName $User.UserPrincipalName

            # Check the MFA status
            $mfaStatus = if ($MsolUser.StrongAuthenticationRequirements) { 
                $MsolUser.StrongAuthenticationRequirements[0].State 
            } else { 
                'Not Enabled' 
            }
             
            $ReportLineAzureAD = [PSCustomObject][Ordered]@{ 
                Name                               = $User.DisplayName
                User                               = $User.Id
                UPN                                = $User.UserPrincipalName
                'Is Licensed'                      = if($User.AssignedLicenses) {'Yes'} else {'No'}
                'Sign in Status'                   = $UserSignInStatus       
                'MFA status'                       = $mfaStatus
                'Days since sign in'               = '0 Not available'
            }
            $Report.Add($ReportLineAzureAD)

            $MFAEnforcedUsers = $Report | Where-Object {$_.'MFA status' -eq 'Enforced'}
            $MFAEnableddUsers =$Report | Where-Object {$_.'MFA status' -eq 'Enabled'}
            $MFANonEnabledUsers =$Report | Where-Object {$_.'MFA status' -eq 'Not Enabled'}
            $MFAEnforcedUsersCount = $MFAEnforcedUsers.count
            $MFANonEnabledUsersCount = $MFANonEnabledUsers.count
            $MFAEnabledUsersCount = $MFAEnabledUsers.count
            
           
            $MFAStats = [PSCustomObject][Ordered]@{
           
                   'Number of user accounts analyzed'              = $Users.Count
                   'Number of accounts Enabled for MFA'            = $MFAEnableddUsers.count
                   'Number of accounts not enabled for MFA'        = $MFANonEnabledUsers.count
                   'Number of accounts with MFA Enforced'          = $MFAEnforcedUsers.count
                   }
                   $MFASummary = $MFAStats
        }
    }

    if (-not $catchBlockExecuted) {
$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($User in $Users) {
    $Authentication = "No sign-in records found"
    $Status = $null; $MFARecordDateTime = $null; $MFAMethodsUsed = $null; $MFAStatus = $null
    $UserLastSignInDate = $null; $DaysSinceLastSignIn = "N/A"
    $UserSignInStatus = if ($User.AccountEnabled -eq $true) {"Allowed"} else {"Blocked"}

    If (!([string]::IsNullOrWhiteSpace($User.signInActivity.lastSignInDateTime))) {
        [datetime]$LastSignIn = $User.signInActivity.lastSignInDateTime
        $DaysSinceLastSignIn = (New-TimeSpan $LastSignIn).Days
    }
    # Get MFA status for the user
    $UserMFAStatus  = $MFAData | Where-Object {$_.Id -eq $User.Id}
    $AuthenticationTypesOutput = $UserMFAStatus.MethodsRegistered -join ", "
    [array]$UserAuditRecords = $AuditRecords | Where-Object {$_.UserId -eq $User.Id} | `
        Sort-Object {$_.CreatedDateTIme -as [datetime]} 
    
    If ($UserAuditRecords) {
        $MFAFlag = $false
        If ("multifactorauthentication" -in $UserAuditRecords.AuthenticationRequirement) {
            # The set of sign-in records contain at least one MFA record, so we extract details
            $MFAUsers++
            $Authentication = "MFA"
            ForEach ($Record in $UserAuditRecords) {
                $Status = $Record.Status.AdditionalDetails
                $MFARecordDateTime = $Record.CreatedDateTIme 
                If ($Status -eq 'MFA completed in Azure AD') {
                    # Found a record that specifies the methods used, so capture that for the report
                    $MFAStatus = "MFA Performed"
                    $MFAMethodsUsed =  $Record.AuthenticationDetails.AuthenticationMethod -join ", "
                    $MFAFlag = $true
                } ElseIf ($MFAFlag -eq $false) {
                    # Otherwise capture details for use of an existing claim
                    $MFAStatus = "Existing claim in the token used"
                    $MFAMethodsUsed = 'Existing claim'                  
                }
            }
        } Else {
            # No MFA sign-in records exist for the user, so they use single-factor
            $Authentication = "Single factor"
        }
    }
    $UserLastSignInDate = $User.SignInActivity.LastSignInDateTime
    
    $ReportLine = [PSCustomObject][Ordered]@{ 
        Name                               = $User.DisplayName
        User                               = $User.Id
        UPN                                = $User.UserPrincipalName
        'Is Licensed'                      = if($User.AssignedLicenses) {'Yes'} else {'No'}
        'Sign in Status'                   = $UserSignInStatus
        LastSignIn                         = $UserLastSignInDate
        'Days since sign in'               = $DaysSinceLastSignIn       
        'Admin flag'                       = $UserMFAStatus.isAdmin 
        'MFA capable'                      = $UserMFAStatus.IsMfaCapable
        'MFA registered'                   = $UserMFAStatus.IsMfaRegistered
         Authentication                    = $Authentication
        'MFA last used'                    = $MFARecordDateTime
        'MFA status'                       = $MFAStatus
        'MFA methods'                      = $MFAMethodsUsed
        'Authentication types'             = $AuthenticationTypesOutput
        'Secondary auth. method'           = $UserMFAStatus.UserPreferredMethodForSecondaryAuthentication
    }
    $Report.Add($ReportLine)

 $AdminUsers = $Report | Where-Object {$_.'Admin Flag' -eq $True}
 $AdminNoMfA = $AdminUsers | Where-Object {$_.'MFA Registered' -eq $False}
 $MFAActiveUsers = $Report | Where-Object {$_.'MFA last used' -ne $null}
 $MFARegisteredUsers =$Report | Where-Object {$_.'MFA Registered' -eq $True}
 $MFANonRegisteredUsers =$Report | Where-Object {$_.'MFA Registered' -eq $False}
 $InactiveLicensedUsers = $Report | Where-Object { $_.'Is Licensed' -eq 'Yes' -and $_.'Days since sign in' -gt 180 -and $_.'Sign in Status' -eq 'Allowed'}
 $UserMFARegisteredCount = ($Report | Where-Object {$_.'MFA Registered' -eq $True}).Count
 $UserMFANotRegisteredCount = ($Report | Where-Object {$_.'MFA Registered' -eq $False}).Count

 $MFAStats = [PSCustomObject][Ordered]@{

        'Admin Users with MFA'                             = $AdminUsers.count
        'Admin Users with no MFA'                          = $AdminNoMFA.Count
        'Number of user accounts analyzed'                 = $Users.Count
        'Number of accounts registered for MFA'            = $MFARegisteredUsers.count
        'Number of accounts not registered for MFA'        = $MFANonRegisteredUsers.count
        'Number of active MFA users in the last 30 days'   = $MFAActiveUsers.count
        'Number of inactive licensed users'                = $InactiveLicensedUsers.count
        }
        $MFASummary = $MFAStats
 }
}
 $EntraReport = $Report 

Write-Host "Finding Admin Users..."
#Admin Users

$AdminUserRoles = New-Object System.Collections.Generic.List[object]
$AllUsers = Get-MgBetaUser -All
$AllRoles = Get-MgDirectoryRole
$AllRoles | ForEach-Object {
    $Role = $_
    $RoleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $Role.Id
    $RoleMembers | ForEach-Object {
        $Member = $_
        $User = $AllUsers | Where-Object { $_.Id -eq $Member.Id }
        if ($User -ne $null) {
            $IsLicensed = if ($User.AssignedLicenses -ne $null) {"Yes"} else {"No"}
            $SignInStatus = if ($User.AccountEnabled -eq $true) {"Allowed"} else {"Blocked"}
            $Result = New-Object PSObject -Property ([ordered]@{
                'DisplayName' = $User.DisplayName
                'UserPrincipalName' = $User.UserPrincipalName
                'Role' = $Role.DisplayName
                'IsLicensed' = $IsLicensed
                'SignInStatus' = $SignInStatus
            })
            $AdminUserRoles.Add($Result)
        }
    }
}
Write-Host "Finding Guest Users..."
#Guest Users

$GuestUsers = New-Object System.Collections.Generic.List[PSCustomObject]
Get-MgBetaUser -All -Filter "UserType eq 'Guest'" -ExpandProperty MemberOf  | ForEach-Object {
    $DisplayName = $_.DisplayName
    $AccountAge = (New-TimeSpan -Start $_.CreatedDateTime).Days
    $Company = $_.CompanyName
    if($Company -eq $null)
    {
        $Company = "-"
    }
    $GroupMembership = @($_.MemberOf.AdditionalProperties.displayName) -join ','
    if($GroupMembership -eq $null)
    {
        $GroupMembership = '-'
    }
    #Add result to array 
    $GuestUsers.Add([PSCustomObject] @{'DisplayName'=$DisplayName;'UserPrincipalName'=$_.UserPrincipalName;'Company'=$Company;'EmailAddress'=$_.Mail;'CreationTime'=$_.CreatedDateTime ;'AccountAge(days)'=$AccountAge;'CreationType'=$_.CreationType;'InvitationAccepted'=$_.ExternalUserState;'GroupMembership'=$GroupMembership})    
}
Write-Host "Collecting Conditional Access Policies..."
# Collect CA Policy
$CAPolicy = Get-MgIdentityConditionalAccessPolicy -All
if (-not $CAPolicy) {
    Write-Host "No CA policies found. Stopping script." -ForegroundColor Red   
}

$CAExport = [PSCustomObject]@()
$CAP = @()
$AdUsers = @()
$Apps = @()
# Extract Values

foreach ( $Policy in $CAPolicy) {
    ### Conditions ###
    $IncludeUG = $null
    $IncludeUG = $Policy.Conditions.Users.IncludeUsers
    $IncludeUG += $Policy.Conditions.Users.IncludeGroups
    $IncludeUG += $Policy.Conditions.Users.IncludeRoles

    $ExcludeUG = $null
    $ExcludeUG = $Policy.Conditions.Users.ExcludeUsers
    $ExcludeUG += $Policy.Conditions.Users.ExcludeGroups
    $ExcludeUG += $Policy.Conditions.Users.ExcludeRoles

    $Apps += $Policy.Conditions.Applications.IncludeApplications
    $Apps += $Policy.Conditions.Applications.ExcludeApplications

    $AdUsers += $ExcludeUG
    $AdUsers += $IncludeUG

    $InclLocation = $null
    $ExclLocation = $null
    $InclLocation = $Policy.Conditions.Locations.includelocations
    $ExclLocation = $Policy.Conditions.Locations.Excludelocations

    $InclPlat = $null
    $ExclPlat = $null
    $InclPlat = $Policy.Conditions.Platforms.IncludePlatforms
    $ExclPlat = $Policy.Conditions.Platforms.ExcludePlatforms
    $InclDev = $null
    $ExclDev = $null
    $InclDev = $Policy.Conditions.Devices.IncludeDevices
    $ExclDev = $Policy.Conditions.Devices.ExcludeDevices
    $devFilters = $null
    $devFilters = $Policy.Conditions.Devices.DeviceFilter.Rule

    $CAExport += New-Object PSObject -Property @{
        ### Users ###
        Users                                               = ""
        Name                                                = $Policy.DisplayName;
        PolicyID                                            = $Policy.ID
        Status                                              = $Policy.State;
        UsersInclude                                        = ($IncludeUG -join ", `r`n");
        UsersExclude                                        = ($ExcludeUG -join ", `r`n");
        ### Cloud apps or actions ###
        'Cloud apps or actions'                             = "";
        ApplicationsIncluded                                = ($Policy.Conditions.Applications.IncludeApplications -join ", `r`n");
        ApplicationsExcluded                                = ($Policy.Conditions.Applications.ExcludeApplications -join ", `r`n");
        userActions                                         = ($Policy.Conditions.Applications.IncludeUserActions -join ", `r`n");
        AuthContext                                         = ($Policy.Conditions.Applications.IncludeAuthenticationContextClassReferences -join ", `r`n");
        ### Conditions ###
        Conditions                                          = "";
        UserRisk                                            = ($Policy.Conditions.UserRiskLevels -join ", `r`n");
        SignInRisk                                          = ($Policy.Conditions.SignInRiskLevels -join ", `r`n");
        PlatformsInclude                                    = ($InclPlat -join ", `r`n");
        PlatformsExclude                                    = ($ExclPlat -join ", `r`n");
        LocationsIncluded                                   = ($InclLocation -join ", `r`n");
        LocationsExcluded                                   = ($ExclLocation -join ", `r`n");
        ClientApps                                          = ($Policy.Conditions.ClientAppTypes -join ", `r`n");
        DevicesIncluded                                     = ($InclDev -join ", `r`n");
        DevicesExcluded                                     = ($ExclDev -join ", `r`n");
        DeviceFilters                                       = ($devFilters -join ", `r`n");

        ### Grant Controls ###
        GrantControls                                       = "";
        BuiltInControls                                     = $($Policy.GrantControls.BuiltInControls)
        TermsOfUse                                          = $($Policy.GrantControls.TermsOfUse)
        CustomControls                                      = $($Policy.GrantControls.CustomAuthenticationFactors)
        GrantOperator                                       = $Policy.GrantControls.Operator

        ### Session Controls ###
        SessionControls                                     = ""
        SessionControlsAdditionalProperties                 = $Policy.SessionControls.AdditionalProperties
        ApplicationEnforcedRestrictionsIsEnabled            = $Policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
        ApplicationEnforcedRestrictionsAdditionalProperties = $Policy.SessionControls.ApplicationEnforcedRestrictions.AdditionalProperties
        CloudAppSecurityType                                = $Policy.SessionControls.CloudAppSecurity.CloudAppSecurityType
        CloudAppSecurityIsEnabled                           = $Policy.SessionControls.CloudAppSecurity.IsEnabled
        CloudAppSecurityAdditionalProperties                = $Policy.SessionControls.CloudAppSecurity.AdditionalProperties
        DisableResilienceDefaults                           = $Policy.SessionControls.DisableResilienceDefaults
        PersistentBrowserIsEnabled                          = $Policy.SessionControls.PersistentBrowser.IsEnabled
        PersistentBrowserMode                               = $Policy.SessionControls.PersistentBrowser.Mode
        PersistentBrowserAdditionalProperties               = $Policy.SessionControls.PersistentBrowser.AdditionalProperties
        SignInFrequencyAuthenticationType                   = $Policy.SessionControls.SignInFrequency.AuthenticationType
        SignInFrequencyInterval                             = $Policy.SessionControls.SignInFrequency.FrequencyInterval
        SignInFrequencyIsEnabled                            = $Policy.SessionControls.SignInFrequency.IsEnabled
        SignInFrequencyType                                 = $Policy.SessionControls.SignInFrequency.Type
        SignInFrequencyValue                                = $Policy.SessionControls.SignInFrequency.Value
        SignInFrequencyAdditionalProperties                 = $Policy.SessionControls.SignInFrequency.AdditionalProperties


    }
}

# Swith user/group Guid to display names

# Filter out Objects
$cajson = $CAExport | ConvertTo-Json -Depth 4
$ADsearch = $AdUsers | Where-Object { $_ -ne 'All' -and $_ -ne 'GuestsOrExternalUsers' -and $_ -ne 'None' }
$AdNames = @{}
Get-MgDirectoryObjectById -ids $ADsearch | ForEach-Object {
    $obj = $_.Id
    #$disp = $_.displayName
    $disp = $_.AdditionalProperties.displayName
    $AdNames.$obj = $disp
    $cajson = $cajson -replace "$obj", "$disp"
}
$CAExport = $cajson | ConvertFrom-Json
# Switch Apps Guid with Display names
$allApps = Get-MgServicePrincipal -All
$allApps | Where-Object { $_.AppId -in $Apps } | ForEach-Object {
    $obj = $_.AppId
    $disp = $_.DisplayName
    $cajson = $cajson -replace "$obj", "$disp"
}
# Switch named location Guid for Display Names
Get-MgIdentityConditionalAccessNamedLocation | ForEach-Object {
    $obj = $_.Id
    $disp = $_.DisplayName
    $cajson = $cajson -replace "$obj", "$disp"
}
# Switch Roles Guid to Names
#Get-MgDirectoryRole | ForEach-Object{
Get-MgDirectoryRoleTemplate | ForEach-Object {
    $obj = $_.Id
    $disp = $_.DisplayName
    $cajson = $cajson -replace "$obj", "$disp"
}
$CAExport = $cajson | ConvertFrom-Json

# Export Setup

$pivot = New-Object System.Collections.Generic.List[PSObject]
$rowItem = New-Object PSObject
$rowitem | Add-Member -type NoteProperty -Name 'CA Item' -Value "row1"
$Pcount = 1
foreach ($CA in $CAExport) {
    $rowitem | Add-Member -type NoteProperty -Name "Policy $pcount" -Value "row1"
    $pcount += 1
}
$pivot.Add($rowItem)

# Add Data to Report
$Rows = $CAExport | Get-Member | Where-Object { $_.MemberType -eq "NoteProperty" }
$Rows | ForEach-Object {
    $rowItem = New-Object PSObject
    $rowname = $_.Name
    $rowitem | Add-Member -type NoteProperty -Name 'CA Item' -Value $_.Name
    $Pcount = 1
    foreach ($CA in $CAExport) {
        $ca | Get-Member | Where-Object { $_.MemberType -eq "NoteProperty" } | ForEach-Object {
            $a = $_.name
            $b = $ca.$a
            if ($a -eq $rowname) {
                $rowitem | Add-Member -type NoteProperty -Name "Policy $pcount" -Value $b
            }
        }
        $pcount += 1
    }
    $pivot.Add($rowItem)
}

# Column Sorting Order
$sort = "Name", "PolicyID", "Status", "Users", "UsersInclude", "UsersExclude", "Cloud apps or actions", "ApplicationsIncluded", "ApplicationsExcluded", `
    "userActions", "AuthContext", "Conditions", "UserRisk", "SignInRisk", "PlatformsInclude", "PlatformsExclude", "LocationsIncluded", `
    "LocationsExcluded", "ClientApps", "Devices", "DevicesIncluded", "DevicesExcluded", "DeviceFilters", `
    "GrantControls", "BuiltInControls", "TermsOfUse", "CustomControls", "GrantOperator", `
    "SessionControls", "SessionControlsAdditionalProperties", "ApplicationEnforcedRestrictionsIsEnabled", "ApplicationEnforcedRestrictionsAdditionalProperties", `
    "CloudAppSecurityType", "CloudAppSecurityIsEnabled", "CloudAppSecurityAdditionalProperties", "DisableResilienceDefaults", "PersistentBrowserIsEnabled", `
    "PersistentBrowserMode", "PersistentBrowserAdditionalProperties", "SignInFrequencyAuthenticationType", "SignInFrequencyInterval", "SignInFrequencyIsEnabled", `
    "SignInFrequencyType", "SignInFrequencyValue", "SignInFrequencyAdditionalProperties"

    $CAP += $pivot  | Where-Object { $_."CA Item" -ne 'row1' } | Sort-object { $sort.IndexOf($_."CA Item") }

    Function ConvertFrom-Html
    {
        [CmdletBinding(SupportsShouldProcess = $True)]
        Param(
            [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true, HelpMessage = "HTML als String")]
            [AllowEmptyString()]
            [string]$Html
        )
    
        if ($PSCmdlet.ShouldProcess("HTML string", "Convert to plaintext"))
        {
            try
            {
                $HtmlObject = New-Object -Com "HTMLFile"
                $HtmlObject.IHTMLDocument2_write($Html)
                $PlainText = $HtmlObject.documentElement.innerText
            }
            catch
            {
                $nl = [System.Environment]::NewLine
                $PlainText = $Html -replace '<br>',$nl
                $PlainText = $PlainText -replace '<br/>',$nl
                $PlainText = $PlainText -replace '<br />',$nl
                $PlainText = $PlainText -replace '</p>',$nl
                $PlainText = $PlainText -replace '<.*?>',''
                $PlainText = [System.Web.HttpUtility]::HtmlDecode($PlainText)
            }
    
            return $PlainText
        }
    }

Write-Host "Retrieving Secure Score..."    
#Secure Score REport

$scores = $null

$Profiles = Get-MgSecuritySecureScoreControlProfile
$Scores = Get-MgSecuritySecureScore 

$latestScore = $scores[0]

$SecureScoreloop = foreach ($Control in $latestScore.controlscores ){
            $controlProfile = $profiles | Where-Object {$_.id -contains $control.controlname}                  
        [PSCustomObject]@{
        'Title'                        = $controlProfile.Title
        'Assessment'                   = $control.description          | ConvertFrom-Html
        'Remediation'                  = $controlProfile.Remediation   | ConvertFrom-Html
        'Reference URL'                = $controlProfile.ActionURL
        'Catagory'                     = $controlProfile.ControlCategory
        'State'                        = $controlProfile.ControlStateUpdates.state | Out-String
        'Comment'                      = $controlProfile.ControlStateUpdates.Comment | Out-String
        'Score'                        = $control.score -replace ",","."
        'Max Score'                    = $controlProfile.MaxScore
        'Additional Points Available'  = $controlProfile.MaxScore - $control.score -replace ",","."
        }       
}

$SecureScore = $SecureScoreloop | Where-Object {$_.'Title' -ne $null} | Select-Object 'Title', 'Assessment', 'Remediation', 'Reference URL', 'Catagory', 'State', 'Comment', 'Score', 'Max Score', 'Additional Points Available'
$SecoreScoreSummary = $latestScore | Select-Object ActiveUserCount, CurrentScore, MaxScore
$SecureScorepercentage = [Math]::Round((($SecoreScoreSummary.CurrentScore / $SecoreScoreSummary.MaxScore) * 100), 2)

$ComparitivesecureScores = Get-MgSecuritySecureScore -Top 1

$allTenantsScore = $ComparitivesecureScores.AverageComparativeScores  | Where-Object { $_.Basis -eq 'AllTenants' }
$totalSeatsScore = $ComparitivesecureScores.AverageComparativeScores  | Where-Object { $_.Basis -eq 'TotalSeats' }

Write-Host "Retrieving Usage Stats..."
# M365 Usage Summary Stats

$ExchangeUsageURI = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageStorage(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $ExchangeUsageURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $ExchangeUsageCSVPath  
$ProgressPreference = 'Continue'
$ExchangeUsageReport = import-csv $ExchangeUsageCSVPath
$ExchangeUsageReport.Foreach({
    $_."storage used (Byte)" = [math]::round($_."storage used (Byte)" / 1GB)
})
 $ExReport = $ExchangeUsageReport | Select-Object @{N = 'Date' ; E = {$_.{Report Date}}}, @{N = 'StorageUsed' ; E = {$_.{storage used (Byte)}}} 

$OneDriveUsageURI = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageStorage(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $OneDriveUsageURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $OneDriveUsageCSVPath
$ProgressPreference = 'Continue'
$OneDriveUsageReport = import-csv $OneDriveUsageCSVPath
$OneDriveUsageReport.Foreach({
    $_."storage used (Byte)" = [math]::round($_."storage used (Byte)" / 1GB)
})
 $ODReport = $OneDriveUsageReport | Where-Object {$_.'Site Type' -eq 'OneDrive'} | Select-Object @{N = 'Date' ; E = {$_.{Report Date}}}, @{N = 'ODStorageUsed' ; E = {$_.{storage used (Byte)}}}

$SharePointUsageURI = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageStorage(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $SharePointUsageURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $SharePointUsageCSVPath
$ProgressPreference = 'Continue' 
$SharePointUsageReport = import-csv $SharePointUsageCSVPath
$SharePointUsageReport.Foreach({
    $_."storage used (Byte)" = [math]::round($_."storage used (Byte)" / 1GB)
})
 $SPReport = $SharePointUsageReport | Select-Object @{N = 'Date' ; E = {$_.{Report Date}}}, @{N = 'SPStorageUsed' ; E = {$_.{storage used (Byte)}}}

 Write-Host "Retrieving User Details..."
 # M365 Usage Details

$UserDetailURI = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $UserDetailURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $UserDetailCSVPath 
$ProgressPreference = 'Continue'
$UserDetailReport = import-csv $UserDetailCSVPath
$USReport = $UserDetailReport |Select-Object 'User Principal Name','Display Name','Is Deleted','Deleted Date','Has Exchange License','Has OneDrive License','Has SharePoint License','Has Teams License','Exchange Last Activity Date','OneDrive Last Activity Date','SharePoint Last Activity Date','Teams Last Activity Date','Exchange License Assign Date','OneDrive License Assign Date','SharePoint License Assign Date','Teams License Assign Date','Assigned Products'

$users = $USReport
$daysAgo180 = (Get-Date).AddDays(-180).ToString("yyyy-MM-dd")
$inactiveUsers = $users | Where-Object {
    # Parse the LastActivityDate as a DateTime object for each service
    $exchangeLastActivityDate = if ($_."Exchange Last Activity Date") { $_."Exchange Last Activity Date" } else { $null }
    $oneDriveLastActivityDate = if ($_."One Drive Last Activity Date") { $_."One Drive Last Activity Date" } else { $null }
    $sharePointLastActivityDate = if ($_."SharePoint Last Activity Date") { $_."SharePoint Last Activity Date" } else { $null }
    $teamsLastActivityDate = if ($_."Teams Last Activity Date") { $_."Teams Last Activity Date" } else { $null }

    # Check if the user is licensed and last had activity more than 180 days ago, or if the user is licensed and has no last activity date
    $_.'Has Exchange License' -eq 'True' -and ($exchangeLastActivityDate -lt $daysAgo180 -or $exchangeLastActivityDate -eq $null) -or
    $_.'Has OneDrive License' -eq 'True' -and ($oneDriveLastActivityDate -lt $daysAgo180 -or $oneDriveLastActivityDate -eq $null) -or
    $_.'Has SharePoint License' -eq 'True' -and ($sharePointLastActivityDate -lt $daysAgo180 -or $sharePointLastActivityDate -eq $null) -or
    $_.'Has Teams License' -eq 'True' -and ($teamsLastActivityDate -lt $daysAgo180 -or $teamsLastActivityDate -eq $null)
}
$InactiveLicensedUsersReport = $inactiveUsers | Select-Object 'User Principal Name', 'Display Name', 'Assigned Products', 'Exchange Last Activity Date', 'OneDrive Last Activity Date', 'SharePoint Last Activity Date', 'Teams Last Activity Date'

$MailboxDetailURI = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $MailboxDetailURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $MailboxDetailCSVPath
$ProgressPreference = 'Continue'
$MailboxDetailReport = import-csv $MailboxDetailCSVPath
$MailboxDetailReport.Foreach({
    $_."storage used (Byte)" = [math]::round($_."storage used (Byte)" / 1GB)
    $_."Issue Warning Quota (Byte)" = [math]::round($_."Issue Warning Quota (Byte)" / 1GB)
    $_."Prohibit Send Quota (Byte)" = [math]::round($_."Prohibit Send Quota (Byte)" / 1GB)
    $_."Prohibit Send/Receive Quota (Byte)" = [math]::round($_."Prohibit Send/Receive Quota (Byte)" / 1GB)
    $_."Deleted Item Size (Byte)" = [math]::round($_."Deleted Item Size (Byte)" / 1GB)
    $_."Deleted Item Quota (Byte)" = [math]::round($_."Deleted Item Quota (Byte)" / 1GB)
    $_ | Add-Member -NotePropertyName "Mailbox Space Remaining (%)" -NotePropertyValue ([math]::round((($_."Prohibit Send Quota (Byte)" - $_."Storage Used (Byte)") / $_."Prohibit Send Quota (Byte)") * 100))
})
 
$O365GroupsDetailURI = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $O365GroupsDetailURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $O365GroupsActivityCSVPath
$ProgressPreference = 'Continue'
$O365GroupsReport = import-csv $O365GroupsActivityCSVPath
$O365GroupsReport.Foreach({
    $_."Exchange Mailbox Storage Used (Byte)" = [math]::round($_."Exchange Mailbox Storage Used (Byte)" / 1GB)
    $_."SharePoint Site Storage Used (Byte)" = [math]::round($_."SharePoint Site Storage Used (Byte)" / 1GB)
})

$OneDriveDetailURI = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $OneDriveDetailURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $OneDriveDetailCSVPath 
$ProgressPreference = 'Continue'
$OneDriveDetailReport = import-csv $OneDriveDetailCSVPath
$OneDriveDetailReport.Foreach({
    $_."Storage Used (Byte)" = [math]::round($_."Storage Used (Byte)" / 1GB)
    $_."Storage Allocated (Byte)" = [math]::round($_."Storage Allocated (Byte)" / 1GB)
    $_ | Add-Member -NotePropertyName "OneDrive Space Remaining (%)" -NotePropertyValue ([math]::round((($_."Storage Allocated (Byte)" - $_."Storage Used (Byte)") / $_."Storage Allocated (Byte)") * 100))
})

$OneDriveActivityURI = "https://graph.microsoft.com/v1.0/reports/getOneDriveActivityUserDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $OneDriveActivityURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $OneDriveActivityCSVPath
$ProgressPreference = 'Continue'
$OneDriveActivityReport = import-csv $OneDriveActivityCSVPath

$SharePointDetailURI = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $SharePointDetailURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $SharePointDetailCSVPath
$ProgressPreference = 'Continue'
$SharePointDetailReport = import-csv $SharePointDetailCSVPath
$SharePointDetailReport.Foreach({
    $_."storage used (Byte)" = [math]::round($_."storage used (Byte)" / 1GB)
    $_."Storage Allocated (Byte)" = [math]::round($_."Storage Allocated (Byte)" / 1GB)
})

$SharePointActivityURI = "https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $SharePointActivityURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $SharePointActivityCSVPath
$ProgressPreference = 'Continue'
$SharePointActivityReport = import-csv $SharePointActivityCSVPath

$TeamsActivityURI = "https://graph.microsoft.com/v1.0/reports/getTeamsTeamActivityDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $TeamsActivityURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $TeamsActivityCSVPath
$ProgressPreference = 'Continue'
$TeamsActivityReport = import-csv $TeamsActivityCSVPath

$TeamsUserActivityURI = "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $TeamsUserActivityURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $TeamsUserActivityCSVPath
$ProgressPreference = 'Continue'
$TeamsUserActivityReport = import-csv $TeamsUserActivityCSVPath

$O365ActivationUserURI = "https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $O365ActivationUserURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $O365ActivationUserCSVPath
$ProgressPreference = 'Continue'
$O365ActivationUserReport = import-csv $O365ActivationUserCSVPath

$M365AppUserDetailURI = "https://graph.microsoft.com/v1.0/reports/getM365AppUserDetail(period='D30')"
$ProgressPreference = 'SilentlyContinue'
Invoke-MgGraphRequest -Uri $M365AppUserDetailURI -Headers $Headers -Method Get -ContentType "application/octet-stream" -OutputFilePath $M365AppUserDetailCSVPath
$ProgressPreference = 'Continue'
$M365AppUserDetailReport = import-csv $M365AppUserDetailCSVPath



Write-Host "Retrieving License Details..."
# M365 Licensing

$licenseSkus = Get-MgSubscribedSku
$translationTable = Invoke-RestMethod -Method Get -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" | ConvertFrom-Csv
$License = New-Object System.Collections.Generic.List[psobject]
foreach ($sku in $licenseSkus) {

    $skuDetails = New-Object -TypeName psobject
    $skuNamePretty = ($translationTable | Where-Object {$_.GUID -eq $sku.skuId} | Sort-Object Product_Display_Name -Unique).Product_Display_Name
    if ($null -eq $skuNamePretty) {
        $skuNamePretty = $sku.SkuPartNumber
    }
    $skuDetails | Add-Member -MemberType NoteProperty -Name "LicenseName" -Value $skuNamePretty
    $skuDetails | Add-Member -MemberType NoteProperty -Name "Quantity" -Value $sku.prepaidUnits.enabled
    $skuDetails | Add-Member -MemberType NoteProperty -Name "Consumed" -Value $sku.consumedunits
    $UnusedLicenses = $sku.prepaidUnits.enabled - $sku.consumedunits
    $skuDetails | Add-Member -MemberType NoteProperty -Name "Unused Licenses" -Value $UnusedLicenses

    $License.Add($skuDetails)
}
$Licenses = $License | Where-Object {$_.LicenseName -ne $null} 

Write-Host "Loading Entra ID Application and Credential Types Report..."
# Entra ID Application and Credential Types Report
$CheckDate = Get-Date
# Define the warning period to check for app secrets that are about to expire
[int]$ExpirationWarningPeriod = 30
# Find registered Entra ID apps that are limited to our organization (not multi-organization)
[array]$RegisteredApps = Get-MgApplication -All -Property Id, appId, displayName, keyCredentials, passwordCredentials, signInAudience | Sort-Object DisplayName
# Remove SharePoint helper apps https://learn.microsoft.com/en-us/answers/questions/1187017/sharepoint-online-client-extensibility-web-applica
$RegisteredApps = $RegisteredApps | Where-Object DisplayName -notLike "SharePoint Online Client Extensibility Web Application Principal*"

If (!($RegisteredApps)) {
    Write-Host "Can't retrieve details of any Entra ID registered apps - exiting"
    
} Else {
    Write-Host ("{0} registered applications found - proceeeding to analyze app secrets" -f $RegisteredApps.count)
}

$AppReport = [System.Collections.Generic.List[Object]]::new() 

ForEach ($App in $RegisteredApps) {
    Write-Host ("Processing {0} app" -f $App.DisplayName)
    $AppOwnersOutput = "No app owner registered"
    # Check for application owners
    [array]$AppOwners = Get-MgApplicationOwner -ApplicationId $App.Id
    If ($AppOwners) {
        $AppOwnersOutput = $AppOwners.additionalProperties.displayName -join ", "
    }

    # Get the app secrets (if any are defined for the app)
[array]$AppSecrets = $App.passwordCredentials

# Check if any secrets exist
if ($AppSecrets.Count -eq 0) {
    # Record that no secrets were found
    $DataLineSecrets = [PSCustomObject] @{
        "Entra ID App Name"   = $App.DisplayName
        "Entra App Id"        = $App.Id
        Owners                = $AppOwnersOutput
        "Credential name"     = "No secrets found"
        "Created"             = $null
        "Credential Id"       = $null
        "Expiration"          = $null
        "Days Until Expiry"   = $null
        Status                = $null
        RecordType            = "Secret"
    }
    $AppReport.Add($DataLineSecrets)
} else {
    ForEach ($AppSecret in $AppSecrets) {
        $ExpirationDays = $null; $Status = $null
        If ($null -ne $AppSecret.endDateTime) {
            $ExpirationDays = (New-TimeSpan -Start $CheckDate -End $AppSecret.endDateTime).Days
            # Figure out app secret status based on the number of days until it expires
            If ($ExpirationDays -lt 0) {
                $Status = "Expired"
            } ElseIf ($ExpirationDays -gt 0 -and $ExpirationDays -le $ExpirationWarningPeriod) {
                $Status = "Expiring soon"
            } Else {
                $Status = "Active"
            }
            # Record what we found
            $DataLineSecrets = [PSCustomObject] @{
                "Entra ID App Name"   = $App.DisplayName
                "Entra App Id"        = $App.Id
                Owners                = $AppOwnersOutput
                "Credential name"     = $AppSecret.DisplayName
                "Created"             = $AppSecret.startDateTime
                "Credential Id"       = $AppSecret.KeyId
                "Expiration"          = $AppSecret.endDateTime
                "Days Until Expiry"   = $ExpirationDays
                Status                = $Status
                RecordType            = "Secret"
            }
            $AppReport.Add($DataLineSecrets)
        }
    }
}

   # Process certificates
[array]$Certificates = $App.keyCredentials

# Check if any certificates exist
if ($Certificates.Count -eq 0) {
    # Record that no certificates were found
    $DataLineCerts = [PSCustomObject] @{
        "Entra ID App Name" = $App.DisplayName
        "Entra App Id"      = $App.Id
        Owners              = $AppOwnersOutput
        "Credential name"   = "No certificates found"
        "Created"           = $null
        "Credential Id"     = $null
        "Expiration"        = $null
        "Days Until Expiry" = $null
        Status              = $null
        RecordType          = "Certificate"
        "Certificate type"  = $null
    }
    $AppReport.Add($DataLineCerts)
} else {
    ForEach ($Certificate in $Certificates) {
        $ExpirationDays = $null; $Status = $null
        If ($null -ne $Certificate.endDateTime) {
            $ExpirationDays = (New-TimeSpan -Start $CheckDate -End $Certificate.endDateTime).Days
            # Figure out app secret status based on the number of days until it expires
            If ($ExpirationDays -lt 0) {
                $Status = "Expired"
            } ElseIf ($ExpirationDays -gt 0 -and $ExpirationDays -le $ExpirationWarningPeriod) {
                $Status = "Expiring soon"
            } Else {
                $Status = "Active"
            }
            # Record what we found
            $DataLineCerts = [PSCustomObject] @{
                "Entra ID App Name" = $App.DisplayName
                "Entra App Id"      = $App.Id
                Owners              = $AppOwnersOutput
                "Credential name"   = $Certificate.DisplayName
                "Created"           = $Certificate.StartDateTime
                "Credential Id"     = $Certificate.KeyId
                "Expiration"        = $Certificate.endDateTime
                "Days Until Expiry" = $ExpirationDays
                Status              = $Status
                RecordType          = "Certificate"
                "Certificate type"  = $Certificate.type
            }
            $AppReport.Add($DataLineCerts)
        }
    }
}
}
$AppReport = $AppReport | Sort-Object RecordType, "Entra ID App Name"

# Application Permisisons Report Section
$EntraApps = Get-MgApplication -All
$permissions = @()

$EntraApps | ForEach-Object {
    $app = $_
    $allRolePermissions = @()
    $allScopePermissions = @()
    $app.RequiredResourceAccess | ForEach-Object {
        $resource = $_
        $rid = $resource.ResourceAppId
        $resourceSP = Get-MgServicePrincipal -Filter "AppId eq '$rid'"
        $resource.ResourceAccess | ForEach-Object {
            $permission = $_
            if ($permission.Type -eq 'Role') {
                $appRoleInfo = $resourceSP.AppRoles | Where-Object Id -eq $permission.Id
                $allRolePermissions += $appRoleInfo.Value 
            }
            elseif ($permission.Type -eq 'Scope') {
                $scopes = $resourceSP.Oauth2PermissionScopes | Where-Object Id -eq $permission.Id
                $allScopePermissions += $scopes.Value 
            }
        }
    }
    $permissions += [PSCustomObject] @{
        "Entra ID App Name" = $app.DisplayName
        "Entra App Id" = $app.AppId
        "Application Permissions" = $allRolePermissions -join ', '
        "Delegated Permissions" = $allScopePermissions -join ', '
        "SignInAudience" = $app.SignInAudience
    }
}

Write-Host "Retrieving DNS Records..."
#Tenant and Domain Report
$Tenant = Get-MgOrganization 
$Domains = $Tenant.VerifiedDomains | Select-Object Name, IsDefault, IsInitial, Type, Capabilities
$TenantProperties = $Tenant | Select-Object DisplayName, CountryLetterCode, Id, OnPremisesSyncEnabled, OnPremisesLastSyncDateTime, TenantType

# Define the DKIM selector
$selector = 'selector1'  # Replace with your actual selector

# Initialize an array to store the DNS records for each domain
$dnsRecords = New-Object System.Collections.Generic.List[PSCustomObject]

# Loop through each domain
foreach ($domain in $Domains.Name) {
    # Get the MX records
    try {
        $mxRecords = Resolve-DnsName -Name $domain -Type MX -ErrorAction Stop
        $mxRecordsString = ($mxRecords | ForEach-Object { "$($_.NameExchange) ($($_.Preference))" }) -join ', '
    } catch {
        $mxRecordsString = "No MX records found for this domain."
    }

    # Get the SPF records
    try {
        $spfRecords = Resolve-DnsName -Name $domain -Type TXT -ErrorAction Stop | Where-Object { $_.Strings -like '*spf*' }
        $spfRecordsString = ($spfRecords | ForEach-Object { $_.Strings -join ' ' }) -join ', '
    } catch {
        $spfRecordsString = "No SPF records found for this domain."
    }

    # Get the DKIM records
    try {
        $dkimRecords = Resolve-DnsName -Name "$selector._domainkey.$domain" -Type TXT -ErrorAction Stop
        $dkimRecordsString = ($dkimRecords | ForEach-Object { $_.Strings -join ' ' }) -join ', '
    } catch {
        $dkimRecordsString = "The specified domain does not have a M365 DKIM key, please check if another provider (E.G Mimecast) is signing email for this domain."
    }

    # Get the DMARC records
    try {
        $dmarcRecords = Resolve-DnsName -Name "_dmarc.$domain" -Type TXT -ErrorAction Stop
        $dmarcRecordsString = ($dmarcRecords | ForEach-Object { $_.Strings -join ' ' }) -join ', '
    } catch {
        $dmarcRecordsString = "No DMARC Records found for this domain."
    }

    $record = [PSCustomObject]@{
        Domain = $domain
        'MX Records' = $mxRecordsString
        'SPF Records' = $spfRecordsString
        'DKIM Records' = $dkimRecordsString
        'DMARC Records' = $dmarcRecordsString
    }
    $dnsRecords.Add($record)
}
Write-Host "Adding Data to HTML Report..."
# Add data to HTML report using PSWriteHTML
 
New-HTML  {
    New-HTMLHeader  {
        New-HTMLSection -Invisible {
            New-HTMLPanel -Invisible {
                New-HTMLImage -Source $MSPLogo -Width 200 -Height 100
            } -AlignContentText right
        }
    }
    New-HTMLTabOptions -BackgroundColor DarkGrey -TextColor White 
    New-HTMLTAB -Name 'M365 Usage Summary' -IconBrands microsoft {
        New-HTMLSection  -CanCollapse -HeaderText 'MFA Summary' -HeaderBackGroundColor '#0b224c' {
            New-HTMLSection -CanCollapse -HeaderText 'MFA Admin Status Chart' -HeaderBackGroundColor '#0b224c' {           
                New-HTMLChart {
	                New-ChartPie -Name 'Admin MFA Registered' -Value ($Report | Where-Object {$_.'Admin Flag' -eq $True}).Count -Color '#6dbf88'
	                New-ChartPie -Name 'Admin MFA not Registered' -Value ($AdminUsers | Where-Object {$_.'MFA Registered' -eq $False}).Count -Color '#db4d2a'
                }
            }	     
            
           
            New-HTMLSection -CanCollapse -HeaderText 'MFA User Status Chart' -HeaderBackGroundColor '#0b224c' {           
                New-HTMLChart {
                    if ($UserMFARegisteredCount -gt 0 -or $UserMFANotRegisteredCount -gt 0) {
                        New-ChartPie -Name 'User MFA Registered' -Value $UserMFARegisteredCount -Color '#6dbf88'
                        New-ChartPie -Name 'User MFA not Registered' -Value $UserMFANotRegisteredCount -Color '#db4d2a'
                    }
                    if ($MFAEnforcedUsersCount -gt 0 -or $MFANonEnabledUsersCount -gt 0 -or $MFAEnabledUsersCount -gt 0) {
                        New-ChartPie -Name 'Number of accounts with MFA Enforced' -Value $MFAEnforcedUsers.count -Color '#6dbf88' 
	                    New-ChartPie -Name 'Number of accounts not enabled for MFA' -Value $MFANonEnabledUsers.count -Color '#db4d2a'  
                        New-ChartPie -Name 'Number of accounts Enabled for MFA' -Value $MFAEnableddUsers.count -Color '#fcaa4c'     
                    }        
                }       
            }
        }
        New-HTMLSection  -CanCollapse -HeaderText 'Mailbox & OneDrive Health' -HeaderBackGroundColor '#0b224c' {
            New-HTMLSection -CanCollapse -HeaderText 'Mailbox Health Chart' -HeaderBackGroundColor '#0b224c' {               
                New-HTMLChart {
                    New-ChartPie -Name 'Mailboxes Healthy' -Value ($MailboxDetailReport | Where-Object {$_."Mailbox Space Remaining (%)" -gt 5}).Count -Color '#6dbf88'
                    New-ChartPie -Name 'Mailboxes Low on Space' -Value ($MailboxDetailReport | Where-Object {$_."Mailbox Space Remaining (%)" -lt 5}).Count -Color '#db4d2a'
                }
            }                        
            New-HTMLSection -CanCollapse -HeaderText 'OneDrive Health Chart' -HeaderBackGroundColor '#0b224c' {               
                New-HTMLChart {
                    New-ChartPie -Name 'OneDrive Accounts Healthy' -Value ($OneDriveDetailReport | Where-Object {$_."OneDrive Space Remaining (%)" -gt 5}).Count -Color '#6dbf88'
                    New-ChartPie -Name 'OneDrive Accounts Low on Space' -Value ($OneDriveDetailReport | Where-Object {$_."OneDrive Space Remaining (%)" -lt 5}).count -Color '#db4d2a'
                        
                }       
            }
        }
        New-HTMLSection -CanCollapse -HeaderText 'Secure Score Summary Chart' -HeaderBackGroundColor '#0b224c' {
            New-HTMLChart {
                    New-ChartLegend -Names 'Current Score', 'Average Tenant Score', 'Average Tenant Score based on seat count' -LegendPosition bottom -Color '#6dbf88', Astral, Saffron
                        New-ChartBar -Name 'Secure Score' -Value $SecureScorepercentage, $allTenantsScore.AverageScore , $totalSeatsScore.AverageScore          
            }       
        }       
         New-HTMLSection  -CanCollapse -HeaderText 'M365 Usage Summary' -HeaderBackGroundColor '#0b224c' {            
            New-HTMLChart -Title 'M365 Usage Graph' -TitleAlignment center {
                New-ChartAxisX -Type datetime -Names $ExReport.Date 
                New-ChartAxisY -TitleText 'Storage Used' -Show -ForceNiceScale 
	            New-ChartLine -Name 'Exchange Storage Used in Gb' -Value $ExReport.storageused -Color Blue 
                New-ChartLine -Name 'OneDrive Storage Used in Gb' -Value $ODReport.ODstorageused -Color Green
                New-ChartLine -Name 'Sharepoint Storage Used in Gb' -Value $SPReport.SPstorageused -Color Purple
            }              	                
        }
    }
    New-HTMLTAB -Name 'Tenant Info' -IconBrands microsoft {
        New-HTMLSection  -CanCollapse -HeaderText 'M365 Tenant Info' -HeaderBackGroundColor '#0b224c' {
            New-HTMLPanel{
                New-HTMLTable -DataTable $TenantProperties   
            }                                                       
        }
        New-HTMLSection -CanCollapse -HeaderText 'M365 Domains' -HeaderBackGroundColor '#0b224c' {
            New-HTMLPanel {
                New-HTMLTable -DataTable $Domains                             
            }
        }
        New-HTMLSection  -CanCollapse -HeaderText 'M365 Domain DNS Info' -HeaderBackGroundColor '#0b224c' {
            New-HTMLPanel{
                New-HTMLTable -DataTable $dnsRecords   
            }                                                       
        }
    }
    New-HTMLTAB -Name 'Licensing' -IconBrands microsoft {
        New-HTMLSection  -CanCollapse -HeaderText 'M365 License Details' -HeaderBackGroundColor '#0b224c' {
            New-HTMLPanel{
                New-HTMLTable -DataTable $Licenses
            }                                	                
        }
    }
    New-HTMLTAB -Name 'Entra ID' -IconBrands microsoft {
        New-HTMLTab -Name 'Entra Signins' {
            New-HTMLSection  -CanCollapse -HeaderText 'Entra ID Sign and MFA Report' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $MFASummary {
                        New-HTMLTableCondition -Name 'Number of accounts not registered for MFA' -ComparisonType number -Operator ne -Value 0 -Color White -BackgroundColor '#db4d2a' -Inline
                        New-HTMLTableCondition -Name 'Admin Users with no MFA' -ComparisonType number -Operator ne -Value 0 -Color White -BackgroundColor '#db4d2a' -Inline
                        New-HTMLTableCondition -Name 'Number of inactive licensed users' -ComparisonType number -Operator ne -Value 0 -Color White -BackgroundColor Orange -Inline
                        New-HTMLTableCondition -Name 'Number of accounts not enabled for MFA' -ComparisonType number -Operator ne -Value 0 -Color White -BackgroundColor '#db4d2a' -Inline
                    }
                }
            }
            New-HTMLSection  -CanCollapse -HeaderText 'Entra ID User Report' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $EntraReport {
                        New-HTMLTableCondition -Name 'Admin flag' -ComparisonType string -Operator eq -Value True -Color White -BackgroundColor Orange -Inline                   
                        New-HTMLTableCondition -Name 'MFA Registered' -ComparisonType string -Operator eq -Value False -Color White -BackgroundColor '#db4d2a' -Inline
                        New-TableConditionGroup -Logic AND {
                        New-HTMLTableCondition -Name 'Days since sign in' -ComparisonType number -Operator gt -Value 180 
                        New-HTMLTableCondition -Name 'Days since sign in' -ComparisonType string -Operator ne -Value 'N/A'
                        } -Color White -BackgroundColor '#db4d2a' -Inline
                        New-TableConditionGroup -Logic AND {
                            New-HTMLTableCondition -Name 'Days since sign in' -ComparisonType number -Operator gt -Value 180 
                            New-HTMLTableCondition -Name 'Is Licensed' -ComparisonType string -Operator eq -Value 'Yes'
                            New-HTMLTableCondition -Name 'Sign in Status' -ComparisonType string -Operator eq -Value 'Allowed'
                            } -BackgroundColor '#fcaa4c' -Inline -Row
                        New-TableConditionGroup -Logic AND {
                            New-HTMLTableCondition -Name 'Sign in Status' -ComparisonType string -Operator eq -Value 'Allowed' 
                            New-HTMLTableCondition -Name 'MFA Status' -ComparisonType string -Operator eq -Value 'Not enabled'
                            } -BackgroundColor '#fcaa4c' -Inline -Row
                    } 
                }
            }
        }
        New-HTMLTAB -Name 'Admin Users & Roles' {
            New-HTMLSection -CanCollapse -HeaderBackGroundColor '#0b224c' { 
                New-HTMLPanel {
                    New-HTMLTable -DataTable $AdminUserRoles
                }
            }
        }
        New-HTMLTAB -Name 'Guest Users' {
            New-HTMLSection -CanCollapse -HeaderBackGroundColor '#0b224c' { 
                New-HTMLPanel {
                    New-HTMLTable -DataTable $GuestUsers
                }
            }
        }
        New-HTMLTAB -Name 'Entra ID CA Policies' {
            New-HTMLSection  -CanCollapse -HeaderText 'Entra ID Conditional Access Policies' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $CAP 
                }
            }
        }    
        New-HTMLTAB -Name 'Entra ID Applications' {              
            New-HTMLSection  -CanCollapse -HeaderText 'Entra ID Applications' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $AppReport {
                        New-HTMLTableCondition -Name 'Status' -ComparisonType string -Operator eq -Value Expired -Color White -BackgroundColor '#db4d2a' -WarningAction SilentlyContinue
                    }                                         
                }
            }                               
            New-HTMLSection -CanCollapse -HeaderText 'Entra ID Application Permissions' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $permissions                              
                }
            }
        } 
    }                  
    New-HTMLTAB -Name 'Secure Score' -IconBrands microsoft {
        New-HTMLSection  -CanCollapse -HeaderText 'SecureScore Report' -HeaderBackGroundColor '#0b224c' {
            New-HTMLPanel {
                New-HTMLTable -DataTable $SecoreScoreSummary
            }
        }
        New-HTMLSection  -CanCollapse -HeaderText 'SecureScore Report' -HeaderBackGroundColor '#0b224c' {
            New-HTMLPanel {
                New-HTMLTable -DataTable $SecureScore {
                    New-HTMLTableCondition -Name 'Additional Points Available' -ComparisonType number -Operator eq -Value 0 -BackgroundColor '#6dbf88' -Row
                    New-HTMLTableCondition -Name 'Additional Points Available' -ComparisonType number -Operator gt -Value 0 -BackgroundColor '#fcaa4c' -Row
                    New-TableConditionGroup -Logic AND {
                    New-HTMLTableCondition -Name 'Score' -ComparisonType number -Operator gt -Value 0     
                    New-HTMLTableCondition -Name 'Additional Points Available' -ComparisonType number -Operator gt -Value 0
                    } -BackgroundColor LightBlue -Row
                }
            }
        }
    }
    New-HTMLTAB -Name 'Usage & Activity Reports' -IconBrands microsoft {   
        New-HTMLTAB -Name 'Users' -IconBrands microsoft {
            New-HTMLSection  -CanCollapse -HeaderText 'M365 User Details' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $USReport
                } 	                
            }
            New-HTMLSection  -CanCollapse -HeaderText 'M365 Inactive User Details' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $InactiveLicensedUsersReport
                } 	                
            }
        }
        New-HTMLTAB -Name 'Mailboxes and Groups' -IconBrands microsoft { 
            New-HTMLSection  -CanCollapse -HeaderText 'M365 User Mailbox Details' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $MailboxDetailReport{
                        New-HTMLTableCondition -Name "Mailbox Space Remaining (%)" -ComparisonType number -Operator lt -value 5  -BackgroundColor '#db4d2a'
                    }                                	                
                }
            }   
            New-HTMLSection  -CanCollapse -HeaderText 'M365 Group Details' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $O365GroupsReport
                }                                	                
            }
        }   
        New-HTMLTAB -Name 'OneDrive' -IconBrands microsoft {
            New-HTMLSection  -CanCollapse -HeaderText 'M365 OneDrive Details' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $OneDriveDetailReport{
                        New-HTMLTableCondition -Name "OneDrive Space Remaining (%)" -ComparisonType number -Operator lt -value 5  -BackgroundColor '#db4d2a'
                    }                 
                }                 
            }
            New-HTMLSection  -CanCollapse -HeaderText 'M365 OneDrive User Activity' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $OneDriveActivityReport{
                        New-HTMLTableCondition -Name 'Shared Externally File Count' -ComparisonType number -Operator gt -Value 0 -BackgroundColor Orange
                    }
                }                 
            }
        }
        New-HTMLTAB -Name 'SharePoint' -IconBrands microsoft { 
            New-HTMLSection  -CanCollapse -HeaderText 'M365 SharePoint Details' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $SharePointDetailReport
                }                 
            }
            New-HTMLSection  -CanCollapse -HeaderText 'M365 SharePoint User Activity' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $SharePointActivityReport{
                        New-HTMLTableCondition -Name 'Shared Externally File Count' -ComparisonType number -Operator gt -Value 0 -BackgroundColor Orange
                    }
                }                 
            }
        }     
        New-HTMLTAB -Name 'Teams' -IconBrands microsoft { 
            New-HTMLSection  -CanCollapse -HeaderText 'M365 Teams Activity' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $TeamsActivityReport
                }                 
            }
            New-HTMLSection  -CanCollapse -HeaderText 'M365 Teams User Activity' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $TeamsUserActivityReport                
                }
            }  
        } 
        New-HTMLTAB -Name 'Apps' -IconBrands microsoft { 
            New-HTMLSection  -CanCollapse -HeaderText 'M365 Activation User Detail' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $O365ActivationUserReport
                }                 
            }
            New-HTMLSection  -CanCollapse -HeaderText 'M365 App User Detail' -HeaderBackGroundColor '#0b224c' {
                New-HTMLPanel {
                    New-HTMLTable -DataTable $M365AppUserDetailReport               
                }
            }  
        } 
    }           
} -FilePath $ReportFile -ShowHTML -FavIcon $FavIcon


Write-Host 'Disconnect Graph session'
Disconnect-MgGraph

$absolutePath = Resolve-Path $ReportFile
Write-Host "The output file is located at: $absolutePath" -ForegroundColor Green



                    
