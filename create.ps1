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
    userPrincipalName = $p.Accounts.AzureAD.userPrincipalName
    phoneNumber = $p.Contact.Personal.Phone.Mobile;
    # One of three:  alternateMobile, office, mobile
    phoneType= 'mobile';
};
$enableSMSSignIn = $false

$aRef = $account.userPrincipalName

if(-Not($dryRun -eq $True)) {
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
        $authenticationMethodSet = $false;
        if( !([string]::IsNullOrEmpty(($getPhoneAuthenticationMethodResponseValue | Out-String))) ){
            if( $getPhoneAuthenticationMethodResponseValue.phoneType.contains($($account.phoneType)) ){
                $authenticationMethodSet = $true;
            }
        }

        switch ($($account.phoneType)){
            # Microsoft docs: https://docs.microsoft.com/nl-nl/graph/api/phoneauthenticationmethod-get?view=graph-rest-beta&tabs=http
            # b6332ec1-7057-4abe-9331-3d72feddfe41 - where phoneType is alternateMobile.
            'alternateMobile' {$phoneTypeId = 'b6332ec1-7057-4abe-9331-3d72feddfe41'}
            # e37fc753-ff3b-4958-9484-eaa9425c82bc - where phoneType is office.
            'office' {$phoneTypeId = 'e37fc753-ff3b-4958-9484-eaa9425c82bc'}
            # 3179e48a-750b-4051-897c-87b9720928f7 - where phoneType is mobile.
            'mobile' {$phoneTypeId = '3179e48a-750b-4051-897c-87b9720928f7'}
        }

        if($authenticationMethodSet -eq $false){
            Write-Verbose -Verbose "No Phone Authentication set for Method $($account.phoneType). Adding Phone Authentication Method $($account.phoneType) : $($account.phoneNumber) for $($account.userPrincipalName).."
        
            $baseUri = "https://graph.microsoft.com/"
            $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods"
        
            $body = @{
                "phoneNumber" = $($account.phoneNumber)
                "phoneType" = $($account.phoneType)
            }
            $bodyJson = $body | ConvertTo-Json -Depth 10
        
            $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Post -Headers $authorization -Body $bodyJson -Verbose:$false
        
            Write-Verbose -Verbose "Successfully Added Phone Authentication Method $($account.phoneType) : $($account.phoneNumber) for $($account.userPrincipalName)"        
        }else{
            $currentPhoneNumber = ($getPhoneAuthenticationMethodResponseValue | Where-Object {$_.phoneType -eq $($account.phoneType)}).phoneNumber
            Write-Verbose -Verbose "Updating current Phone Authentication Method $($account.phoneType) : $currentPhoneNumber for $($account.userPrincipalName).."

            $previousAccount = [PSCustomObject]@{
                userPrincipalName = $account.userPrincipalName
                phoneNumber = $currentPhoneNumber;
                # One of three:  alternateMobile, office, mobile
                phoneType= $account.phoneType;
            }

            $baseUri = "https://graph.microsoft.com/"
            $addPhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods/$phoneTypeId"
        
            $body = @{
                "phoneNumber" = $($account.phoneNumber)
                "phoneType" = $($account.phoneType)
            }
            $bodyJson = $body | ConvertTo-Json -Depth 10
        
            $addPhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $addPhoneAuthenticationMethodUri -Method Put -Headers $authorization -Body $bodyJson -Verbose:$false
        
            Write-Verbose -Verbose "Successfully updated Phone Authentication Method $($account.phoneType) : $($account.phoneNumber) for $($account.userPrincipalName)"
        }

        if($enableSMSSignIn -eq $true){
            Write-Verbose -Verbose "Enabling Phone Authentication Method for $($account.userPrincipalName).."

            $baseUri = "https://graph.microsoft.com/"
            $enablePhoneAuthenticationMethodUri = $baseUri + "/beta/users/$($account.userPrincipalName)/authentication/phoneMethods/$phoneTypeId/enableSmsSignIn"

            $enablePhoneAuthenticationMethodResponse = Invoke-RestMethod -Uri $enablePhoneAuthenticationMethodUri -Method Post -Headers $authorization -Verbose:$false

            Write-Verbose -Verbose "Successfully enabled Phone Authentication Method for $($account.userPrincipalName)"
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
        userPrincipalName = $account.userPrincipalName;
        phoneNumber = $account.phoneNumber;
        phoneType = $account.phoneType;
    };
};

Write-Output $result | ConvertTo-Json -Depth 10;