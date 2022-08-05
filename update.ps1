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
    userPrincipalName               = $p.Accounts.MicrosoftAzureAD.userPrincipalName
    # Phone numbers use the format "+<country code> <number>x<extension>", with extension optional.
    # For example, +1 5555551234 or +1 5555551234x123 are valid. Numbers are rejected when creating/updating if they do not match the required format.    
    mobile                          = "+31 " + $p.Contact.Business.Phone.Mobile
    onlySetMobileWhenEmpty          = $false
    alternateMobile                 = "+31 " + $p.Contact.Personal.Phone.Mobile
    onlySetAlternateMobileWhenEmpty = $false
    office                          = "+31 " + $p.Contact.Business.Phone.Fixed
    onlySetOfficeWhenEmpty          = $false

    enableSMSSignInMobile           = $false
}

# Troubleshooting
# $dryRun = $false
# $account = [PSCustomObject]@{
#     userPrincipalName               = 'j.doe@enyoi.org'
#     # Phone numbers use the format "+<country code> <number>x<extension>", with extension optional.
#     # For example, +1 5555551234 or +1 5555551234x123 are valid. Numbers are rejected when creating/updating if they do not match the required format.    
#     mobile                          = "+31 " + '0612345678'
#     onlySetMobileWhenEmpty          = $false
#     alternateMobile                 = "+31 " + '0687654321'
#     onlySetAlternateMobileWhenEmpty = $false
#     office                          = "+31 " + '0229123456'
#     onlySetOfficeWhenEmpty          = $false

#     enableSMSSignInMobile           = $false
# }

# Get current Azure AD user and authentication methods 
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
    $searchUri = $baseGraphUri + "v1.0/users/$($account.userPrincipalName)"

    Write-Verbose "Querying Azure AD user with UPN $($account.userPrincipalName)"
    $azureUser = Invoke-RestMethod -Uri $searchUri -Method Get -Headers $authorization -Verbose:$false
    if ($null -eq $azureUser.id) { throw "Could not find Azure user $($account.userPrincipalName)" }
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
            Action  = "UpdateAccount"
            Message = "Error correlating to and updating Azure MFA settings of account with id $($aRef). Error Message: $auditErrorMessage"
            IsError = $True
        })
}

# Set Phone Authentication Method mobile
try {
    if ( ![string]::IsNullOrEmpty($account.mobile) ) {
        # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
        # 3179e48a-750b-4051-897c-87b9720928f7 - where phoneType is mobile.
        $phoneType = "mobile"
        $phoneTypeId = '3179e48a-750b-4051-897c-87b9720928f7'
        $phoneNumber = "$($account.mobile)"

        $authenticationMethodSet = $false
        if ( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ) {
            if ( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($phoneType)) ) {
                $authenticationMethodSet = $true
            }
        }

        if ($authenticationMethodSet -eq $false) {
            Write-Verbose "No Phone Authentication set for Method $($phoneType). Adding Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef)"
            
            $baseUri = "https://graph.microsoft.com/"
            $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods"
            
            $body = @{
                "phoneNumber" = $phoneNumber
                "phoneType"   = $phoneType
            }
            $bodyJson = $body | ConvertTo-Json -Depth 10
            
            if (-Not($dryRun -eq $True)) {
                $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Post -Headers $authorization -Body $bodyJson -Verbose:$false
                
                  
                $auditLogs.Add([PSCustomObject]@{
                        Action  = "UpdateAccount"
                        Message = "Successfully added Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef)"
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: No Phone Authentication set for Method $($phoneType). Adding Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef)"
            }
        }
        else {
            $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object { $_.phoneType -eq $($phoneType) }).phoneNumber           
            if ($account.onlySetMobileWhenEmpty -eq $true) {
                Write-Warning "Phone Authentication Method $($phoneType) set to only update when empty. Since this already contains data ($currentPhoneNumber), skipped update for account with id $($aRef)"
            }
            else {
                Write-Verbose "Updating current Phone Authentication Method $($phoneType) value '$currentPhoneNumber' to value '$($phoneNumber)' for account with id $($aRef)"

                $baseUri = "https://graph.microsoft.com/"
                $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods/$phoneTypeId"
                
                $body = @{
                    "phoneNumber" = $phoneNumber
                    "phoneType"   = $phoneType
                }
                $bodyJson = $body | ConvertTo-Json -Depth 10
                
                if (-Not($dryRun -eq $True)) {
                    $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Put -Headers $authorization -Body $bodyJson -Verbose:$false
                
                      
                    $auditLogs.Add([PSCustomObject]@{
                            Action  = "UpdateAccount"
                            Message = "Successfully updated Phone Authentication Method $($phoneType) value '$currentPhoneNumber' to value '$($phoneNumber)' for account with id $($aRef)"
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Updating current Phone Authentication Method $($phoneType) : $currentPhoneNumber for account with id $($aRef)"
                }
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
            Action  = "UpdateAccount"
            Message = "Error setting Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef). Error message: $($auditErrorMessage)"
            IsError = $True
        })
}

    
# Enable SMS Sign-in
try {
    if ($account.enableSMSSignInMobile -eq $true) {
        Write-Verbose "Enabling $($phoneType) SMS Sign-in for account with id $($aRef)"

        $baseUri = "https://graph.microsoft.com/"
        $enablePhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods/$phoneTypeId/enableSmsSignIn"

        if (-Not($dryRun -eq $True)) {
            $enablePhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $enablePhoneAuthenticationMethodUri -Method Post -Headers $authorization -Verbose:$false

              
            $auditLogs.Add([PSCustomObject]@{
                    Action  = "UpdateAccount"
                    Message = "Successfully enabled $($phoneType) SMS Sign-in for account with id $($aRef)"
                    IsError = $false
                })
        }
        else {
            Write-Warning "DryRun: Enabling $($phoneType) SMS Sign-in for account with id $($aRef)"
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
            Action  = "UpdateAccount"
            Message = "Error enabling $($phoneType) SMS Sign-in for account with id $($aRef). Error message: $($auditErrorMessage)"
            IsError = $True
        })
}

# Set Phone Authentication Method alternateMobile
try {
    if ( ![string]::IsNullOrEmpty($account.alternateMobile) ) {
        # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
        # b6332ec1-7057-4abe-9331-3d72feddfe41 - where phoneType is alternateMobile.
        $phoneType = "alternateMobile"
        $phoneTypeId = 'b6332ec1-7057-4abe-9331-3d72feddfe41'
        $phoneNumber = "$($account.alternateMobile)"

        $authenticationMethodSet = $false
        if ( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ) {
            if ( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($phoneType)) ) {
                $authenticationMethodSet = $true
            }
        }

        if ($authenticationMethodSet -eq $false) {
            Write-Verbose "No Phone Authentication set for Method $($phoneType). Adding Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef)"
                
            $baseUri = "https://graph.microsoft.com/"
            $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods"
                
            $body = @{
                "phoneNumber" = $phoneNumber
                "phoneType"   = $phoneType
            }
            $bodyJson = $body | ConvertTo-Json -Depth 10
                
            if (-Not($dryRun -eq $True)) {
                $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Post -Headers $authorization -Body $bodyJson -Verbose:$false
                
                  
                $auditLogs.Add([PSCustomObject]@{
                        Action  = "UpdateAccount"
                        Message = "Successfully added Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef)"
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: No Phone Authentication set for Method $($phoneType). Adding Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef)"
            }     
        }
        else {
            $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object { $_.phoneType -eq $($phoneType) }).phoneNumber        
            if ($account.onlySetAlternateMobileWhenEmpty -eq $true) {
                Write-Warning "Phone Authentication Method $($phoneType) set to only update when empty. Since this already contains data ($currentPhoneNumber), skipped update for account with id $($aRef)"
            }
            else {
                Write-Verbose "Updating current Phone Authentication Method $($phoneType) : $currentPhoneNumber for account with id $($aRef)"

                $baseUri = "https://graph.microsoft.com/"
                $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods/$phoneTypeId"
                
                $body = @{
                    "phoneNumber" = $phoneNumber
                    "phoneType"   = $phoneType
                }
                $bodyJson = $body | ConvertTo-Json -Depth 10
                
                if (-Not($dryRun -eq $True)) {
                    $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Put -Headers $authorization -Body $bodyJson -Verbose:$false
                
                      
                    $auditLogs.Add([PSCustomObject]@{
                            Action  = "UpdateAccount"
                            Message = "Successfully updated Phone Authentication Method $($phoneType) value '$currentPhoneNumber' to value '$($phoneNumber)' for account with id $($aRef)"
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Updating current Phone Authentication Method $($phoneType) : $currentPhoneNumber for account with id $($aRef)"
                }    
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
            Action  = "UpdateAccount"
            Message = "Error setting Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef). Error message: $($auditErrorMessage)"
            IsError = $True
        })
}

# Set Phone Authentication Method office
try {
    if ( ![string]::IsNullOrEmpty($account.office) ) {
        # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
        # e37fc753-ff3b-4958-9484-eaa9425c82bc - where phoneType is office.
        $phoneType = "office"
        $phoneTypeId = 'e37fc753-ff3b-4958-9484-eaa9425c82bc'
        $phoneNumber = "$($account.office)"

        $authenticationMethodSet = $false
        if ( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ) {
            if ( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($phoneType)) ) {
                $authenticationMethodSet = $true
            }
        }

        if ($authenticationMethodSet -eq $false) {
            Write-Verbose "No Phone Authentication set for Method $($phoneType). Adding Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef)"
                
            $baseUri = "https://graph.microsoft.com/"
            $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods"
                
            $body = @{
                "phoneNumber" = $phoneNumber
                "phoneType"   = $phoneType
            }
            $bodyJson = $body | ConvertTo-Json -Depth 10
                
            if (-Not($dryRun -eq $True)) {
                $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Post -Headers $authorization -Body $bodyJson -Verbose:$false
                
                  
                $auditLogs.Add([PSCustomObject]@{
                        Action  = "UpdateAccount"
                        Message = "Successfully added Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef)"
                        IsError = $false
                    })
            }
            else {
                Write-Warning "DryRun: No Phone Authentication set for Method $($phoneType). Adding Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef)"
            }    
        }
        else {
            $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object { $_.phoneType -eq $($phoneType) }).phoneNumber        
            if ($account.onlySetOfficeWhenEmpty -eq $true) {
                Write-Warning "Phone Authentication Method $($phoneType) set to only update when empty. Since this already contains data ($currentPhoneNumber), skipped update for account with id $($aRef)"
            }
            else {
                Write-Verbose "Updating current Phone Authentication Method $($phoneType) : $currentPhoneNumber for account with id $($aRef)"

                $baseUri = "https://graph.microsoft.com/"
                $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($aRef)/authentication/phoneMethods/$phoneTypeId"
                
                $body = @{
                    "phoneNumber" = $phoneNumber
                    "phoneType"   = $phoneType
                }
                $bodyJson = $body | ConvertTo-Json -Depth 10
                
                if (-Not($dryRun -eq $True)) {
                    $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Put -Headers $authorization -Body $bodyJson -Verbose:$false
                
                      
                    $auditLogs.Add([PSCustomObject]@{
                            Action  = "UpdateAccount"
                            Message = "Successfully updated Phone Authentication Method $($phoneType) value '$currentPhoneNumber' to value '$($phoneNumber)' for account with id $($aRef)"
                            IsError = $false
                        })
                }
                else {
                    Write-Warning "DryRun: Updating current Phone Authentication Method $($phoneType) : $currentPhoneNumber for account with id $($aRef)"
                }    
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
            Action  = "UpdateAccount"
            Message = "Error setting Phone Authentication Method $($phoneType) with value '$($phoneNumber)' for account with id $($aRef). Error message: $($auditErrorMessage)"
            IsError = $True
        })
}

Write-Warning $success

# Send results
$result = [PSCustomObject]@{
    Success          = $success
    AccountReference = $aRef
    AuditLogs        = $auditLogs
    Account          = $account

    # Optionally return data for use in other systems
    ExportData       = [PSCustomObject]@{
        id                    = $azureUser.id
        userPrincipalName     = $azureUser.userPrincipalName
        mobile                = $account.mobile
        alternateMobile       = $account.alternateMobile
        office                = $account.office
        enableSMSSignInMobile = $account.enableSMSSignInMobile
    }
}

Write-Output $result | ConvertTo-Json -Depth 10