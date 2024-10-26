# Sjekker om Microsoft Graph modulen er installert, hvis ikke gjøres dette
$module = Get-Module -Name Microsoft.Graph -ListAvailable

if ($module -eq $null) {
    Write-Host "Microsoft Graph PowerShell module is not installed. Installing now..."

    # Setter PSrepository til PSGallery som trusted
    $trusted = Get-PSRepository -Name PSGallery | Select-Object -ExpandProperty InstallationPolicy
    if ($trusted -ne "Trusted") {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }

    # Installerer modulen
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force

    Write-Host "Microsoft Graph PowerShell module has been installed successfully."
} else {
    Write-Host "Microsoft Graph PowerShell module is already installed."
}

# Kobler til M365 Tenant med Micorsoft Graph PowerShell modul
$TenantID = "cc3d062d-6a68-49a5-82a5-f6bbf7d96507"
Connect-MgGraph -TenantId $TenantID `
    -Scope  "User.ReadWrite.All", `
            "Group.ReadWrite.All", `
            "Directory.ReadWrite.All", `
            "RoleManagement.ReadWrite.Directory"

$Details = Get-MgContext
$Scopes = $Details | Select-Object -ExpandProperty Scopes
$Scopes = $Scopes -join ","
$OrgName = (Get-MgOrganization).DisplayName
""
""
"Microsoft Graph current session details:"
"----------------------------------------"
"Tenant Id = $($Details.TenantId)"
"Client Id = $($Details.ClientId)"
"Org Name  = $OrgName"
"App Name  = $($Details.AppName)"
"Account   = $($Details.Account)"
"Scopes    = $Scopes"
"----------------------------------------"


# Sjekker om ExchangeOnlineManagement modulen er installert, hvis ikke gjøres dette
if(!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module not found. Installing..."
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
} else {
    Write-Host "ExchangeOnlineManagement module is already installed."
}

# Importerer ExchangeOnlineManagement modulen
Import-Module ExchangeOnlineManagement

# Kobler til Exchange Online<
try {
    Connect-ExchangeOnline -ShowProgress $true
    Write-Host "Successfully connected to Exchange Online."
}

catch {
    Write-Host "Error connecting to Exchange Online: $_"
}


# Henter inn brukere fra CSV-filen
$users = Import-CSV -Path 'Users.csv' -Delimiter ","

$PasswordProfile = @{
    Password = '3fs#dsaDAf224s#'
    ForceChangePasswordNextSignIn = $true
}


# Lager brukerene
foreach ($user in $users) {
    $Params = @{
        UserPrincipalName = $user.givenName + "." + $user.surName + "@DigSecIndustry.onmicrosoft.com"
        DisplayName = $user.givenName + " " + $user.surName
        GivenName = $user.GivenName
        SurName = $user.surName
        MailNickname = $user.givenName + "." + $user.surName
        AccountEnabled = $true
        PasswordProfile = $PasswordProfile
        Department = $user.Department
        CompanyName = $user.CompanyName
        Country = $user.Country
        City = $user.City
        JobTitle = $user.JobTitle
    }
    New-MgUser @Params
} 

Write-Host "All users are created"

# Lager Gjestebruker
$User = @{
    GivenName = "Dagfinn"
    SurName = "Haaland"
    Department = "Ekstern"
    CompanyName = "Ekstern"
    Country = "Norway"
    City = "Gjovik"
    JobTitle = "EksternKonsulent"
}
$Params = @{
    UserPrincipalName = $user.givenName + "." + $user.surName + "@DigSecIndustry.onmicrosoft.com"
    DisplayName = $user.givenName + " " + $user.surName
    GivenName = $user.GivenName
    SurName = $user.surName
    MailNickname = $user.givenName + "." + $user.surName
    AccountEnabled = $true
    PasswordProfile = $PasswordProfile
    Department = $user.Department
    CompanyName = $user.CompanyName
    Country = $user.Country
    City = $user.City
    JobTitle = $user.JobTitle
}
New-MgUser @Params

# Lager grupper
$departments = @("Ledelse", "Utvikling", "Salg", "Kundesupport", "IT-drift", "Administrasjon", "Ekstern")
foreach ($department in $departments) {
    $membershiprule = "user.department -eq `"$department`""
    $Params = @{
        DisplayName = "$department"
        Description = "Gruppe for ansatte i avdelingen $department"
        MailEnabled = $true
        MailNickname = $department
        SecurityEnabled = $true
        GroupTypes = @("Unified", "DynamicMembership")
        Membershiprule = $membershiprule
        MembershipRuleProcessingState = "On"
    }
    New-MgGroup @Params
}

Write-Host "All groups are created"

# Lager Securtiy group for lisensiering
$Params = @{
    DisplayName = "m365-E5-license"
    Description = "Members of this group will get a M365 E5 license"
    MailEnabled = $false
    MailNickname = "m365-e5-license"
    SecurityEnabled = $true
    GroupTypes = @("DynamicMembership")
    MembershipRule = '(user.companyName -eq "DigSecIndustry")'
    MembershipRuleProcessingState = "On"
}

New-MgGroup @Params

Write-Host "Secuirty Group for M365 E5 license is created"