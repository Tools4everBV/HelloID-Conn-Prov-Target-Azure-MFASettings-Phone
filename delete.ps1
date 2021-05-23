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
    removeAlternateMobile           = $true;
    removeOffice                    = $true;
    removeMobile                    = $true;
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
        if( $account.removeAlternateMobile -eq $true ){
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
                Write-Verbose -Verbose "No Phone Authentication set. Nothing to delete"
            }else{
                $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object {$_.phoneType -eq $($account.phoneType)}).phoneNumber
                Write-Verbose -Verbose "Deleting current Phone Authentication Method $($account.phoneType) : $currentPhoneNumber for $($account.userPrincipalName).."

                $baseUri = "https://graph.microsoft.com/"
                $deletePhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods/$phoneTypeId"
            
                $deletePhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $deletePhoneAuthenticationMethodUri -Method Delete -Headers $authorization -Body $bodyJson -Verbose:$false
            
                Write-Verbose -Verbose "Successfully deleted Phone Authentication Method $($account.phoneType) : $currentPhoneNumber for $($account.userPrincipalName)"
            }
        }
        
        if( $account.removeOffice -eq $true ){
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
                Write-Verbose -Verbose "No Phone Authentication set. Nothing to delete"
            }else{
                $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object {$_.phoneType -eq $($account.phoneType)}).phoneNumber
                Write-Verbose -Verbose "Deleting current Phone Authentication Method $($account.phoneType) : $currentPhoneNumber for $($account.userPrincipalName).."

                $baseUri = "https://graph.microsoft.com/"
                $deletePhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods/$phoneTypeId"
            
                $deletePhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $deletePhoneAuthenticationMethodUri -Method Delete -Headers $authorization -Body $bodyJson -Verbose:$false
            
                Write-Verbose -Verbose "Successfully deleted Phone Authentication Method $($account.phoneType) : $currentPhoneNumber for $($account.userPrincipalName)"
            }
        }

        if( $account.removeMobile -eq $true ){
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
                Write-Verbose -Verbose "No Phone Authentication set. Nothing to delete"
            }else{
                $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object {$_.phoneType -eq $($account.phoneType)}).phoneNumber
                Write-Verbose -Verbose "Deleting current Phone Authentication Method $($account.phoneType) : $currentPhoneNumber for $($account.userPrincipalName).."

                $baseUri = "https://graph.microsoft.com/"
                $deletePhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods/$phoneTypeId"
            
                $deletePhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $deletePhoneAuthenticationMethodUri -Method Delete -Headers $authorization -Body $bodyJson -Verbose:$false
            
                Write-Verbose -Verbose "Successfully deleted Phone Authentication Method $($account.phoneType) : $currentPhoneNumber for $($account.userPrincipalName)"
            }
        }
    }

    $auditLogs.Add([PSCustomObject]@{
        Action = "DeleteAccount"
        Message = "Updated Azure MFA settings of account with UPN $($aRef)"
        IsError = $false;
    });

    $success = $true;    
}catch{
    $auditLogs.Add([PSCustomObject]@{
        Action = "DeleteAccount"
        Message = "Error updating Azure MFA settings of account with UPN $($aRef): $($_)"
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
        removeMobile            = $account.removeMobile
        removeAlternateMobile   = $account.removeAlternateMobile
        removeOffice            = $account.removeOffice
    };
};

Write-Output $result | ConvertTo-Json -Depth 10;
