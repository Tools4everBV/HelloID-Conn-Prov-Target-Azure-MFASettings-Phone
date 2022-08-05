# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

$VerbosePreference = "SilentlyContinue"
$InformationPreference = "Continue"
$WarningPreference = "Continue"

#Initialize default properties
$p = $person | ConvertFrom-Json
$m = $manager | ConvertFrom-Json
$aRef = $accountReference | ConvertFrom-Json
$mRef = $managerAccountReference | ConvertFrom-Json
$success = $true # Set to true at start, because only when an error occurs it is set to false
$auditLogs = [Collections.Generic.List[PSCustomObject]]::new()

# AzureAD Application Parameters #
$config = ConvertFrom-Json $configuration

$AADtenantID = $config.AADtenantID
$AADAppId = $config.AADAppId
$AADAppSecret = $config.AADAppSecret
# Change mapping here
$account = [PSCustomObject]@{
    removeAlternateMobile = $true
    removeOffice          = $true
    removeMobile          = $true
}

# Troubleshooting
# $dryRun = $false

try {
    Write-Verbose "Generating Microsoft Graph API Access Token"
    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$AADTenantID/oauth2/token"

    $body = @{
        grant_type    = "client_credentials"
        client_id     = "$AADAppId"
        client_secret = "$AADAppSecret"
        resource      = "https://graph.microsoft.com"
    }

    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token

    #Add the authorization header to the request
    $authorization = @{
        Authorization  = "Bearer $accesstoken"
        'Content-Type' = "application/json"
        Accept         = "application/json"
    }

    $baseGraphUri = "https://graph.microsoft.com/"
    $searchUri = $baseGraphUri + "v1.0/users/$($aRef)"

    Write-Verbose "Querying Azure AD user with UPN $($aRef)"
    $azureUser = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    if ($null -eq $azureUser.id) { throw "Could not find Azure user $($aRef)" }
    Write-Verbose "Successfully queried Azure AD user $($azureUser.userPrincipalName) ($($azureUser.id))"
    # Set aRef to use for further actions
    $aRef = $azureUser.id

    Write-Verbose "Gathering current Phone Authentication Methods for account with id $($aRef)"

    $baseUri = "https://graph.microsoft.com/"
    $getPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods"

    $getPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $getPhoneAuthenticationMethodUri -Method Get -Headers $authorization -Verbose:$false
    $getPhoneAuthenticationMethodResponseValue = $getPhoneAuthenticationMethodResponse.value
    Write-Verbose ("Current phone authentication method: " + ($getPhoneAuthenticationMethodResponseValue | Out-String) )
}
catch {
    $ex = $PSItem
    $verboseErrorMessage = $ex
    Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)"

    $auditErrorMessage = ($ex | ConvertFrom-Json).error.message
    $success = $false  
    $auditLogs.Add([PSCustomObject]@{
            Action  = "DeleteAccount"
            Message = "Error correlating to and updating Azure MFA settings of account with id $($aRef). Error Message: $auditErrorMessage"
            IsError = $True
        })
}


# Remove Phone Authentication Method alternateMobile
try {
    if ( $account.removeAlternateMobile -eq $true ) {
        # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
        # b6332ec1-7057-4abe-9331-3d72feddfe41 - where phoneType is alternateMobile.
        $phoneType = "alternateMobile"
        $phoneTypeId = 'b6332ec1-7057-4abe-9331-3d72feddfe41'

        $authenticationMethodSet = $false
        if ( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ) {
            if ( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($phoneType)) ) {
                $authenticationMethodSet = $true
            }
        }

        if ($authenticationMethodSet -eq $false) {
            Write-Warning "No Phone Authentication set for Method $($phoneType) for account with id $($aRef). Nothing to delete"
        }
        else {
            $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object { $_.phoneType -eq $($phoneType) }).phoneNumber
            Write-Verbose "Deleting current Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef)"

            $baseUri = "https://graph.microsoft.com/"
            $deletePhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods/$phoneTypeId"

            if (-Not($dryRun -eq $True)) {
                $deletePhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $deletePhoneAuthenticationMethodUri -Method Delete -Headers $authorization -Body $bodyJson -Verbose:$false

                $auditLogs.Add([PSCustomObject]@{
                        Action  = "DeleteAccount"
                        Message = "Successfully deleted Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef)"
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: Deleting current Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef)"
            }

        }
    }
}
catch {
    $ex = $PSItem
    $verboseErrorMessage = $ex
    Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)"

    $auditErrorMessage = ($ex | ConvertFrom-Json).error.message
    $success = $false  
    $auditLogs.Add([PSCustomObject]@{
            Action  = "DeleteAccount"
            Message = "Error deleting Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef). Error Message: $auditErrorMessage"
            IsError = $True
        })
}

# Remove Phone Authentication Method office
try {
    if ( $account.removeOffice -eq $true ) {
        # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
        # e37fc753-ff3b-4958-9484-eaa9425c82bc - where phoneType is office.
        $phoneType = "office"
        $phoneTypeId = 'e37fc753-ff3b-4958-9484-eaa9425c82bc'

        $authenticationMethodSet = $false
        if ( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ) {
            if ( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($phoneType)) ) {
                $authenticationMethodSet = $true
            }
        }

        if ($authenticationMethodSet -eq $false) {
            Write-Warning "No Phone Authentication set for Method $($phoneType) for account with id $($aRef). Nothing to delete"
        }
        else {
            $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object { $_.phoneType -eq $($phoneType) }).phoneNumber
            Write-Verbose "Deleting current Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef)"

            $baseUri = "https://graph.microsoft.com/"
            $deletePhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods/$phoneTypeId"

            if (-Not($dryRun -eq $True)) {
                $deletePhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $deletePhoneAuthenticationMethodUri -Method Delete -Headers $authorization -Body $bodyJson -Verbose:$false

                $auditLogs.Add([PSCustomObject]@{
                        Action  = "DeleteAccount"
                        Message = "Successfully deleted Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef)"
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: Deleting current Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef)"
            }

        }
    }
}
catch {
    $ex = $PSItem
    $verboseErrorMessage = $ex
    Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)"

    $auditErrorMessage = ($ex | ConvertFrom-Json).error.message
    $success = $false  
    $auditLogs.Add([PSCustomObject]@{
            Action  = "DeleteAccount"
            Message = "Error deleting Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef). Error Message: $auditErrorMessage"
            IsError = $True
        })
}


# Remove Phone Authentication Method mobile
try {
    if ( $account.removeMobile -eq $true ) {
        # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
        # 3179e48a-750b-4051-897c-87b9720928f7 - where phoneType is mobile.
        $phoneType = "mobile"
        $phoneTypeId = '3179e48a-750b-4051-897c-87b9720928f7'

        $authenticationMethodSet = $false
        if ( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ) {
            if ( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($phoneType)) ) {
                $authenticationMethodSet = $true
            }
        }

        if ($authenticationMethodSet -eq $false) {
            Write-Warning "No Phone Authentication set for Method $($phoneType) for account with id $($aRef). Nothing to delete"
        }
        else {
            $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object { $_.phoneType -eq $($phoneType) }).phoneNumber
            Write-Verbose "Deleting current Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef)"

            $baseUri = "https://graph.microsoft.com/"
            $deletePhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods/$phoneTypeId"

            if (-Not($dryRun -eq $True)) {
                $deletePhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $deletePhoneAuthenticationMethodUri -Method Delete -Headers $authorization -Body $bodyJson -Verbose:$false

                $auditLogs.Add([PSCustomObject]@{
                        Action  = "DeleteAccount"
                        Message = "Successfully deleted Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef)"
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: Deleting current Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef)"
            }

        }
    }
}
catch {
    $ex = $PSItem
    $verboseErrorMessage = $ex
    Write-Verbose "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($verboseErrorMessage)"

    $auditErrorMessage = ($ex | ConvertFrom-Json).error.message
    $success = $false  
    $auditLogs.Add([PSCustomObject]@{
            Action  = "DeleteAccount"
            Message = "Error deleting Phone Authentication Method $($phoneType) with value '$($currentPhoneNumber)' for account with id $($aRef). Error Message: $auditErrorMessage"
            IsError = $True
        })
}

# Send results
$result = [PSCustomObject]@{
    Success          = $success
    AccountReference = $aRef
    AuditLogs        = $auditLogs
    Account          = $account
    PreviousAccount  = $previousAccount

    # Optionally return data for use in other systems
    ExportData       = [PSCustomObject]@{
        id                = $azureUser.id
        userPrincipalName = $azureUser.userPrincipalName
        alternateMobile   = $(if ($true -eq $account.removeAlternateMobile) { '' })
        office            = $(if ($true -eq $account.removeOffice) { '' })
        mobile            = $(if ($true -eq $account.removeMobile) { '' })
    }
}

Write-Output $result | ConvertTo-Json -Depth 10