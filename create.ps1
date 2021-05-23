#Initialize default properties
$p = $person | ConvertFrom-Json;
$m = $manager | ConvertFrom-Json;
$aRef = $accountReference | ConvertFrom-Json;
$mRef = $managerAccountReference | ConvertFrom-Json;
$success = $False;
$auditLogs = New-Object Collections.Generic.List[PSCustomObject];

# AzureAD Application Parameters #
$config = ConvertFrom-Json $configuration

$AADtenantID = $config.AADtenantID
$AADAppId = $config.AADAppId
$AADAppSecret = $config.AADAppSecret

# Set TLS to accept TLS, TLS 1.1 and TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls12

# Change mapping here
$account = [PSCustomObject]@{
    userPrincipalName               = $p.Accounts.MicrosoftActiveDirectory.userPrincipalName
    # Phone numbers use the format "+<country code> <number>x<extension>", with extension optional.
    # For example, +1 5555551234 or +1 5555551234x123 are valid. Numbers are rejected when creating/updating if they do not match the required format.    
    mobile                          = "+31" + $p.Contact.Business.Phone.Mobile;
    onlySetMobileWhenEmpty          = $false;
    alternateMobile                 = "+31" + $p.Contact.Personal.Phone.Mobile;
    onlySetAlternateMobileWhenEmpty = $false;
    office                          = "+31" + $p.Contact.Business.Phone.Fixed;
    onlySetOfficeWhenEmpty          = $true;

    enableSMSSignInMobile           = $false
};

try{
    Write-Verbose -Verbose "Generating Microsoft Graph API Access Token.."
    $baseUri = "https://login.microsoftonline.com/"
    $authUri = $baseUri + "$AADTenantID/oauth2/token"

    $body = @{
        grant_type      = "client_credentials"
        client_id       = "$AADAppId"
        client_secret   = "$AADAppSecret"
        resource        = "https://graph.microsoft.com"
    }

    $Response = Invoke-RestMethod -Method POST -Uri $authUri -Body $body -ContentType 'application/x-www-form-urlencoded'
    $accessToken = $Response.access_token;

    #Add the authorization header to the request
    $authorization = @{
        Authorization = "Bearer $accesstoken";
        'Content-Type' = "application/json";
        Accept = "application/json";
    }

    Write-Verbose -Verbose "Gathering current Phone Authentication Methods for $($account.userPrincipalName).."

    $baseUri = "https://graph.microsoft.com/"
    $getPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods"

    $getPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $getPhoneAuthenticationMethodUri -Method Get -Headers $authorization -Verbose:$false
    $getPhoneAuthenticationMethodResponseValue = $getPhoneAuthenticationMethodResponse.value
    Write-Verbose -Verbose ("Phone authentication method: " + ($getPhoneAuthenticationMethodResponseValue | Out-String) )

    if(-Not($dryRun -eq $True)) {
        if( ![string]::IsNullOrEmpty($account.mobile) ){
            # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
            # 3179e48a-750b-4051-897c-87b9720928f7 - where phoneType is mobile.
            $phoneType = "mobile"
            $phoneTypeId = '3179e48a-750b-4051-897c-87b9720928f7'

            $authenticationMethodSet = $false
            if( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ){
                if( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($phoneType)) ){
                    $authenticationMethodSet = $true;
                }
            }

            if($authenticationMethodSet -eq $false){
                Write-Verbose -Verbose "No Phone Authentication set for Method $($phoneType). Adding Phone Authentication Method $($phoneType) : $($account.phoneNumber) for $($account.userPrincipalName).."
            
                $baseUri = "https://graph.microsoft.com/"
                $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods"
            
                $body = @{
                    "phoneNumber" = $($account.mobile)
                    "phoneType" = $($phoneType)
                }
                $bodyJson = $body | ConvertTo-Json -Depth 10
            
                $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Post -Headers $authorization -Body $bodyJson -Verbose:$false
            
                Write-Verbose -Verbose "Successfully Added Phone Authentication Method $($phoneType) : $($account.phoneNumber) for $($account.userPrincipalName)"        
            }else{
                $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object {$_.phoneType -eq $($phoneType)}).phoneNumber           
                if($account.onlySetMobileWhenEmpty -eq $true){
                     Write-Verbose -Verbose "Phone Authentication Method $($phoneType) set to only update when empty. Since this already contains data ($currentPhoneNumber), skipped update for $($account.userPrincipalName)"
                }else{
                    Write-Verbose -Verbose "Updating current Phone Authentication Method $($phoneType) : $currentPhoneNumber for $($account.userPrincipalName).."

                    $baseUri = "https://graph.microsoft.com/"
                    $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods/$phoneTypeId"
                
                    $body = @{
                        "phoneNumber" = $($account.mobile)
                        "phoneType" = $($phoneType)
                    }
                    $bodyJson = $body | ConvertTo-Json -Depth 10
                
                    $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Put -Headers $authorization -Body $bodyJson -Verbose:$false
                
                    Write-Verbose -Verbose "Successfully updated Phone Authentication Method $($phoneType) : $($account.phoneNumber) for $($account.userPrincipalName)"
                }
            }

            if($account.enableSMSSignInMobile -eq $true){
                Write-Verbose -Verbose "Enabling $($phoneType) SMS Sign-in for $($account.userPrincipalName).."

                $baseUri = "https://graph.microsoft.com/"
                $enablePhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods/$phoneTypeId/enableSmsSignIn"

                $enablePhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $enablePhoneAuthenticationMethodUri -Method Post -Headers $authorization -Verbose:$false

                Write-Verbose -Verbose "Successfully enabled $($phoneType) SMS Sign-in for $($account.userPrincipalName)"
            }        
        }

        if( ![string]::IsNullOrEmpty($account.alternateMobile) ){
            # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
            # b6332ec1-7057-4abe-9331-3d72feddfe41 - where phoneType is alternateMobile.
            $phoneType = "alternateMobile"
            $phoneTypeId = 'b6332ec1-7057-4abe-9331-3d72feddfe41'

            $authenticationMethodSet = $false
            if( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ){
                if( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($phoneType)) ){
                    $authenticationMethodSet = $true;
                }
            }

            if($authenticationMethodSet -eq $false){
                Write-Verbose -Verbose "No Phone Authentication set for Method $($phoneType). Adding Phone Authentication Method $($phoneType) : $($account.phoneNumber) for $($account.userPrincipalName).."
            
                $baseUri = "https://graph.microsoft.com/"
                $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods"
            
                $body = @{
                    "phoneNumber" = $($account.alternateMobile)
                    "phoneType" = $($phoneType)
                }
                $bodyJson = $body | ConvertTo-Json -Depth 10
            
                $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Post -Headers $authorization -Body $bodyJson -Verbose:$false
            
                Write-Verbose -Verbose "Successfully Added Phone Authentication Method $($phoneType) : $($account.phoneNumber) for $($account.userPrincipalName)"        
            }else{
                $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object {$_.phoneType -eq $($phoneType)}).phoneNumber        
                if($account.onlySetAlternateMobileWhenEmpty -eq $true){
                    Write-Verbose -Verbose "Phone Authentication Method $($phoneType) set to only update when empty. Since this already contains data ($currentPhoneNumber), skipped update for $($account.userPrincipalName)"
               }else{
                   Write-Verbose -Verbose "Updating current Phone Authentication Method $($phoneType) : $currentPhoneNumber for $($account.userPrincipalName).."

                   $baseUri = "https://graph.microsoft.com/"
                   $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods/$phoneTypeId"
               
                   $body = @{
                       "phoneNumber" = $($account.mobile)
                       "phoneType" = $($phoneType)
                   }
                   $bodyJson = $body | ConvertTo-Json -Depth 10
               
                   $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Put -Headers $authorization -Body $bodyJson -Verbose:$false
               
                   Write-Verbose -Verbose "Successfully updated Phone Authentication Method $($phoneType) : $($account.phoneNumber) for $($account.userPrincipalName)"
               }
            } 
        }    

        if( ![string]::IsNullOrEmpty($account.office) ){
            # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
            # e37fc753-ff3b-4958-9484-eaa9425c82bc - where phoneType is office.
            $phoneType = "office"
            $phoneTypeId = 'e37fc753-ff3b-4958-9484-eaa9425c82bc'

            $authenticationMethodSet = $false
            if( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ){
                if( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($phoneType)) ){
                    $authenticationMethodSet = $true;
                }
            }

            if($authenticationMethodSet -eq $false){
                Write-Verbose -Verbose "No Phone Authentication set for Method $($phoneType). Adding Phone Authentication Method $($phoneType) : $($account.phoneNumber) for $($account.userPrincipalName).."
            
                $baseUri = "https://graph.microsoft.com/"
                $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods"
            
                $body = @{
                    "phoneNumber" = $($account.office)
                    "phoneType" = $($phoneType)
                }
                $bodyJson = $body | ConvertTo-Json -Depth 10
            
                $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Post -Headers $authorization -Body $bodyJson -Verbose:$false
            
                Write-Verbose -Verbose "Successfully Added Phone Authentication Method $($phoneType) : $($account.phoneNumber) for $($account.userPrincipalName)"        
            }else{
                $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object {$_.phoneType -eq $($phoneType)}).phoneNumber        
                if($account.onlySetOfficeWhenEmpty -eq $true){
                    Write-Verbose -Verbose "Phone Authentication Method $($phoneType) set to only update when empty. Since this already contains data ($currentPhoneNumber), skipped update for $($account.userPrincipalName)"
               }else{
                   Write-Verbose -Verbose "Updating current Phone Authentication Method $($phoneType) : $currentPhoneNumber for $($account.userPrincipalName).."

                   $baseUri = "https://graph.microsoft.com/"
                   $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods/$phoneTypeId"
               
                   $body = @{
                       "phoneNumber" = $($account.mobile)
                       "phoneType" = $($phoneType)
                   }
                   $bodyJson = $body | ConvertTo-Json -Depth 10
               
                   $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Put -Headers $authorization -Body $bodyJson -Verbose:$false
               
                   Write-Verbose -Verbose "Successfully updated Phone Authentication Method $($phoneType) : $($account.phoneNumber) for $($account.userPrincipalName)"
               }
            } 
        }
    }

    $auditLogs.Add([PSCustomObject]@{
        Action = "CreateAccount"
        Message = "Correlated to and updated Azure MFA settings of account with UPN $($aRef)"
        IsError = $false;
    });

    $success = $true;    
}catch{
    $auditLogs.Add([PSCustomObject]@{
        Action = "CreateAccount"
        Message = "Error correlating to and updating Azure MFA settings of account with UPN $($aRef): $($_)"
        IsError = $True
    });
    Write-Error $_;
}

# Send results
$result = [PSCustomObject]@{
	Success= $success;
	AccountReference= $aRef;
	AuditLogs = $auditLogs;
    Account = $account;
    PreviousAccount = $previousAccount;

    # Optionally return data for use in other systems
    ExportData = [PSCustomObject]@{
        userPrincipalName       = $account.userPrincipalName;
        mobile                  = $account.mobile
        alternateMobile         = $account.alternateMobile
        office                  = $account.office
        enableSMSSignInMobile   = $account.enableSMSSignInMobile
    };
};

Write-Output $result | ConvertTo-Json -Depth 10;