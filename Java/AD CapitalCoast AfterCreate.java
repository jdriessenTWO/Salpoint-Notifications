###############################################################################################################################
# GENERIC HELPER FUNCTIONS
###############################################################################################################################
#
#
# JD  16/02/2026 - Update
###############################################################################################################################

#Logging
function WriteToLog([String] $info) {
    try {
        #log file info
        $logDate = Get-Date -UFormat "%Y%m%d"
        $logFile = "C:\SailPoint\Logs\AD_AfterCreate_$logDate.log"
        $logEntry = $nativeIdentity + ": " + $info
        $logEntry | Out-File -FilePath $logFile -Append -Force -ErrorAction Stop 2>$null
    } catch {
        # Log failure is ignored, continue execution
    }
}

function Get-AttributeValueFromAccountRequest([sailpoint.Utils.objects.AccountRequest] $request, [String] $targetAttribute) {
    if ($request) {
        foreach ($attrib in $request.AttributeRequests) {
            if ($attrib.Name -eq $targetAttribute) {
                return $attrib.Value
            }
        }
    } else {
        WriteToLog("Account request object was null")
    }
    return $null
}

function Load-ADCredentials {
    try {
        $appObject = [sailpoint.Utils.xml.XmlFactory]::Instance.parseXml($env:Application)
        if ($appObject -and $appObject.domainSettings[0] -and $appObject.domainSettings[0].password) {
            $svcAccPwdDecoded = [sailpoint.Utils.tools.Util]::decode($appObject.domainSettings[0].password, $true)
            return New-Object System.Management.Automation.PSCredential ($appObject.domainSettings[0].user, (ConvertTo-SecureString $svcAccPwdDecoded -AsPlainText -Force))
        }
        return $null
    }
    catch {
        return $null
    }
}

function Connect-ExchangeServer {
    try {
        $appObject = [sailpoint.Utils.xml.XmlFactory]::Instance.parseXml($env:Application)
        if ($appObject -and $appObject.exchangeSettings[0] -and $appObject.exchangeSettings[0].ExchHost[0]) {
            $exchUserPwdDecoded = [sailpoint.Utils.tools.Util]::decode($appObject.exchangeSettings[0].password, $true)
            $exchUserCred = New-Object System.Management.Automation.PSCredential ($appObject.exchangeSettings[0].user, (ConvertTo-SecureString $exchUserPwdDecoded -AsPlainText -Force))
            $exchUri = "http://$($appObject.exchangeSettings[0].ExchHost[0])/powershell/?SerializationLevel=Full"
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $exchUri -Credential $exchUserCred -Authentication Kerberos -ea Stop
            $exchSession = Import-PSSession $Session -AllowClobber -DisableNameChecking
            WriteToLog("Exchange connected")
            return $exchSession
        }
        WriteToLog("Failed to retrieve Exchange server or credentials")
        return $null
    }
    catch {
        WriteToLog("Error connecting to Exchange: $_")
        return $null
    }
}

function Send-UserPassword {
    $appObject = [sailpoint.Utils.xml.XmlFactory]::Instance.parseXml($env:Application)
    $Password = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "password"
    if ($Password) {
        $decryptedPwd = [sailpoint.utils.tools.util]::decode($Password, $true)
    }

    $toRecipients = @()
    if ($reportingManagerEmail -and $reportingManagerEmail -ne "") {
        $toRecipients += [PSCustomObject]@{ emailAddress = @{ address = $reportingManagerEmail } }
    }

    if ($hiringManagerEmail -and $hiringManagerEmail -ne "" -and $hiringManagerEmail -ne $reportingManagerEmail) {
        $toRecipients += [PSCustomObject]@{ emailAddress = @{ address = $hiringManagerEmail } }
    }

    if ($toRecipients.Count -gt 0 -and $decryptedPwd) {
        try {
            # Azure AD app credentials
            $clientId = $appObject.entraEmailClientId
            $clientSecret = $appObject.entraEmailClientSecret
            $tenantId = $appObject.entraEmailTenantId

            # Get access token
            $tokenBody = @{
                client_id     = $clientId
                scope         = "https://graph.microsoft.com/.default"
                client_secret = $clientSecret
                grant_type    = "client_credentials"
            }
            $token = (Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -Body $tokenBody -ContentType "application/x-www-form-urlencoded").access_token

            # Email content
            $htmlBody = @"
    <p><strong>(CONFIDENTIAL) Please do not forward this e-mail</strong></p>
    <p>Kia Ora,</p>
    <p>Here is the password for the account requested for <b>$($adUser.Name)</b>. This account will provide access to Health New Zealand IT systems.</p>
    <div style='border: 1px solid #ccc; padding: 10px; background-color: #f9f9f9; display: inline-block;'>$($decryptedPwd)</div>
    <p>This is a temporary one time use password for <b>$($adUser.Name)</b> to use to log into our Health NZ network when they will be asked to change this to a password of their choosing.</p>
    <p>For security reasons we sent the digital account details to you in a separate email with the subject line <b>'The digital account created for $($adUser.Name)'</b> and we ask that you don't forward on or send this email. Please advise <b>$($adUser.Name)</b> of their login details and password in person or in a Teams/mobile call. As a very last resort you can share these in a SMS but this is not Health NZ preferred method.</p>
    <p>If you need help to complete this, go to <a href='https://support.tewhatuora.govt.nz/esc?id=emp_taxonomy_topic&topic_id=58b3b0d91b450610d6102069b04bcb20'>Kāpehu (ServiceNow)</a> and raise a request for support to our Digital Services Service Desk.</p>
    <br>
    <p>He waka eke noa,</p>
    <p>IT Service Desk</p>
    <p>Service Experience</p>
    <p>Digital Enterprise Services</p>
"@

            $mail = @{
                message = @{
                    subject = "Digital access password for $($adUser.Name)"
                    body = @{
                        contentType = "HTML"
                        content = $htmlBody
                    }
                    toRecipients = $toRecipients
                    from = @{
                        emailAddress = @{
                            address = "Identity.Notification@tewhatuora.govt.nz"
                        }
                    }
                }
                saveToSentItems = "false"
            } | ConvertTo-Json -Depth 10 -Compress

            # Send email
            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/Identity.Notification@tewhatuora.govt.nz/sendMail" -Method Post -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" } -Body $mail
            WriteToLog("Sent initial password to $reportingManagerEmail and $hiringManagerEmail")
        }
        catch {
            WriteToLog("Failed to send initial password email: $_")
        }
    }
}

function Set-HomeDrivePermissions {
    ##### HOME DIRECTORY SETUP
    WriteToLog("Starting homeDirectory setup.")
    $homeDirectory = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "homeDirectory"
    $sAMAccountName = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "sAMAccountName"
    if ($homeDirectory -and $homeDirectory.EndsWith($sAMAccountName)) {
        WriteToLog("Valid home directory found: $homeDirectory")
        $tempDrive = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 8 | ForEach-Object {[char]$_})

        # Map the users home directory with credentials
        $parentHomeDirectory = Split-Path -Path $homeDirectory -Parent
        WriteToLog("Mounting: $($parentHomeDirectory) as $($tempDrive)")
        New-PSDrive -Name $tempDrive -PSProvider FileSystem -Root $parentHomeDirectory -Credential $adCredentials

        # Check if the new users path exists and is a directory
        $homeDirectoryDrive = "$($tempDrive):\$($sAMAccountName)"
        if (Test-Path $homeDirectoryDrive -PathType Container) {
            WriteToLog("Found existing folder: $($tempDrive):\$($sAMAccountName)")
            # If the folder already exists, rename it to avoid conflicts
            $null = Rename-Item -Path $homeDirectoryDrive -NewName ($sAMAccountName + "." + (Get-Date -Format 'yyyyMMdd_HHmm') + ".old")
            WriteToLog("Renamed existing folder: $($tempDrive):\$($sAMAccountName)")
        }

        #Create the new Home directory folder
        $null = New-Item -ItemType Directory -Path $homeDirectoryDrive
        WriteToLog("Created new folder on DFS for: $($homeDirectory)")

        # Check if the new folder was created successfully
        if (Test-Path $homeDirectoryDrive -PathType Container) {
            # Ensure permissions on the destination folder are correct so the User (and Admins) have access
            WriteToLog("Started setting permissions for new homeDirectory: $($homeDirectory)")

            $FolderACL = $null
            $FolderACL = Get-Acl $homeDirectoryDrive -ErrorAction SilentlyContinue
            if (($FolderACL) -and (-not($FolderACL.AreAccessRulesCanonical))) #if we got an ACL and that ACL is non-canonical (incorrectly ordered)
            {
                Set-Acl -Path $homeDirectoryDrive -AclObject $FolderACL
                WriteToLog("Fixing ACLs.")
            }
            elseif (-not($FolderACL)) #if we didn't get the ACL's for the folder
            {
                WriteToLog("Unable to find existing ACLs")
            }

            # Set the new owner using the new account
            WriteToLog("Setting new user as owner.")
            $acl = Get-Acl -Path $homeDirectoryDrive
            $newUser = "AD\$($sAMAccountName)"
            $owner = New-Object System.Security.Principal.NTAccount($newUser)
            $acl.SetOwner($owner)
            Set-Acl -Path $homeDirectoryDrive -AclObject $acl

            # Get System Permissions for the top level folder
            WriteToLog("Adding permissions as required.")
            $acl = Get-Acl -Path $homeDirectoryDrive
            $accountsToAdd = @(
                "AD\$($sAMAccountName)"
            )
            foreach ($account in $accountsToAdd) {
                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $account,
                    "FullControl",
                    "None",
                    "none",
                    "Allow"
                )
                $acl.AddAccessRule($accessRule)
                Set-Acl -Path $homeDirectoryDrive -AclObject $acl
            }
            
                        
            WriteToLog("Adding permissions as required.")
            $acl = Get-Acl -Path $homeDirectoryDrive
            $accountsToAdd = @(
                "AD\NS-LS-UserProfileAdmins"
            )
            foreach ($account in $accountsToAdd) {
                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $account,
                    "FullControl",
                    "ContainerInherit,ObjectInherit",
                    "none",
                    "Allow"
                )
                $acl.AddAccessRule($accessRule)
                Set-Acl -Path $homeDirectoryDrive -AclObject $acl
            }

            WriteToLog("Adding permissions as required.")
            $acl = Get-Acl -Path $homeDirectoryDrive
            $accountsToAdd = @(
                "CREATOR OWNER"
            )
            foreach ($account in $accountsToAdd) {
                $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    $account,
                    "FullControl",
                    "ContainerInherit,ObjectInherit",
                    "InheritOnly",
                    "Allow"
                )
                $acl.AddAccessRule($accessRule)
                Set-Acl -Path $homeDirectoryDrive -AclObject $acl
            }
        
            #Remove access we dont need
            WriteToLog("Removing unwanted ACLs.")
            $acl = Get-Acl -Path $homeDirectoryDrive
            $accountsToRemove = @(
                "BUILTIN\Users",
                "NT AUTHORITY\Authenticated Users"
            )
            foreach ($account in $accountsToRemove) {
                $acl = Get-Acl -Path $homeDirectoryDrive
                $accessRule = $acl.Access | Where-Object { $_.IdentityReference -eq $account }
                if ($accessRule) {
                    $acl.RemoveAccessRule($accessRule)
                }
                Set-Acl -Path $homeDirectoryDrive -AclObject $acl
            }

            # Disable inheritance and remove inherited access rules
            $acl.SetAccessRuleProtection($true, $true)

            # Apply the updated ACL to the directory
            Set-Acl -Path $homeDirectoryDrive -AclObject $acl

            WriteToLog("Finished setting permissions for: $($homeDirectory)")
        } else
        {
            WriteToLog("Unable to find new folder on DFS for: $($homeDirectory)")
        }

        # Clean up by removing the temporary PSDrive
        Remove-PSDrive -Name $tempDrive -ErrorAction SilentlyContinue
        WriteToLog("Un-mounting: $($parentHomeDirectory)")
    }
}

function Add-UserToGroups {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.ActiveDirectory.Management.ADEntity]$ADUser,

        [Parameter(Mandatory = $true)]
        [array]$Groups
    )

    foreach ($GroupName in $Groups) {
        # Try to find the group in AD
        $GroupSearch = Get-ADGroup -Filter { Name -eq $GroupName } -ErrorAction SilentlyContinue
        if ($GroupSearch) {
            try {
                # Attempt to add the user to the group
                WriteToLog("Adding user to group: " + $GroupSearch.DistinguishedName)
                Add-ADGroupMember -Identity $GroupSearch.DistinguishedName -Members $ADUser.DistinguishedName -Server $adServer -Credential $adCredentials -ErrorAction Stop
            }
            catch {
                # Log if theres an error adding the user
                WriteToLog("ERROR - Failed to add user to group: " + $GroupSearch.DistinguishedName + " - " + $_.Exception.Message)
            }
        } else {
            # Log if the group was not found
            WriteToLog("WARNING - Unable to find group: " + $GroupName)
        }
    }
}

###############################################################################################################################
# BODY
###############################################################################################################################

try {
    Add-Type -Path C:\SailPoint\IQService\utils.dll
    Import-Module ActiveDirectory -ErrorAction Stop

    # Request object
    $requestObject = [sailpoint.Utils.objects.AccountRequest]::new([sailpoint.utils.xml.XmlUtil]::getReader([System.IO.StringReader]::new([System.String]$env:Request)))
    $nativeIdentity = $requestObject.nativeIdentity
    if (!$nativeIdentity) { throw "Unable to find new native identity from provisioning. Exiting Script!" } else { WriteToLog("Found Native Identity: $($nativeIdentity)") }

    # Result object
    $resultObject = [sailpoint.Utils.objects.ServiceResult]::new([sailpoint.utils.xml.XmlUtil]::getReader([System.IO.StringReader]::new([System.String]$env:Result)))
   
    # AD Server
    $adServer = $resultObject.Attributes["createdOnServer"]
    if(!$adServer) {$adServer = (Get-ADDomainController).Hostname}
    WriteToLog("Working on AD Server: $adServer")

    # Load AD connector credentials
    $adCredentials = Load-ADCredentials
    if (!$adCredentials) { throw "Unable to get AD credentials to do updates. Exiting Script!" }

    # Load the user from AD
    $adUser = Get-ADUser -Identity $nativeIdentity -Server $adServer -Properties *
    if (!$adUser) { throw "Unable to find new AD User on server. Exiting Script!" } else { WriteToLog("Found AD User: $($adUser.Name)") }

    #Setup variables for later use
    $employeeType = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "employeeType"
    $reportingManagerEmail = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "reportingManagerEmail"
    $hiringManagerEmail = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "hiringManagerEmail"

    $sAMAccountName = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "sAMAccountName"
    $firstName = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "givenName"
    $firstName = ($firstName.Normalize([System.Text.NormalizationForm]::FormD) -replace '\p{Mn}', '')

    $lastName = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "sn"
    $lastName = ($lastName.Normalize([System.Text.NormalizationForm]::FormD) -replace '\p{Mn}', '')

    ### Email setup

    ### Generate base UPN
    $domain = Get-AttributeValueFromAccountRequest -request $requestObject -targetAttribute "domainSuffix"
    $baseUpn = "$($firstName.ToLower()).$($lastName.ToLower())" -replace "\s", "" 
    $baseUpn = $baseUpn -replace "[^A-Za-z0-9._-]", ""
    $baseUpn = $baseUpn -replace "\.+", "."
    $baseUpn = $baseUpn -replace "^[._-]+|[._-]+$", ""

    $prefix = $baseUpn.Substring(0, [Math]::Min(64, $baseUpn.Length))
    $filterProxy = "*$prefix*"
    $upn = "$prefix@$domain"
    $counter = 0
    WriteToLog("Built initial UPN attempt: $upn")

    # Check for uniqueness in both UserPrincipalName and proxyAddresses
    while (Get-ADUser -Filter {(UserPrincipalName -eq $upn) -or (proxyAddresses -like $filterProxy)} -ErrorAction SilentlyContinue) {
        $counter++
        if($counter -lt 10) {
            $prefix = $baseUpn.Substring(0, [Math]::Min(63, $baseUpn.Length)) 
        }
        else {
            $prefix = $baseUpn.Substring(0, [Math]::Min(62, $baseUpn.Length))
        }
        $filterProxy = "*$prefix$counter*"
        $upn = "$prefix$counter@$domain"
    }
    
    # User variables
    $userPrincipalName = $upn
    if (!$userPrincipalName) { throw "Unable to generate new AD User UPN. Exiting Script!" } else { WriteToLog("Found UPN: $userPrincipalName") }

    WriteToLog("User type is: $($employeeType)")
    if (($employeeType -ine "External") -and ($employeeType -ine "Vendor") -and ($employeeType -ine "Contractor") -and ($employeeType -ine "GP SCP") -and ($employeeType -ine "GP Medtech 32"))
    {

        Set-ADUser -Identity $nativeIdentity -Server $adServer -UserPrincipalName $userPrincipalName -Credential $adCredentials

        WriteToLog("Starting mailbox setup logic.")
        WriteToLog("Setting up Office 365 mailbox for: $($adUser.SamAccountName)")

        # Connect to Exchange Online
        $exchSession = Connect-ExchangeServer
        if (!$exchSession) { throw "Unable to connect to Exchange Online. Exiting script!" }

        # Enable remote mailbox (OFFICE365)
        $RandomNumber = Get-Random -Minimum 10 -Maximum 99
        $UPNPrefix = ($userPrincipalName.Split('@')[0])
        $RoutingAddress = "SMTP:$UPNPrefix$RandomNumber@3dhb.mail.onmicrosoft.com"

        Enable-RemoteMailbox -Identity $sAMAccountName -RemoteRoutingAddress $RoutingAddress -DomainController $adServer -ErrorAction Stop
        Enable-RemoteMailbox -Identity $sAMAccountName -Archive -DomainController $adServer -ErrorAction Stop
        Start-Sleep -Seconds 10
        $newMailbox = Get-RemoteMailbox -Identity $sAMAccountName -DomainController $adServer

        # Setup the mail address attribute
        $newEmailAddress = if ($newMailbox -and $newMailbox.PrimarySmtpAddress) { $newMailbox.PrimarySmtpAddress } else { $userPrincipalName }
        WriteToLog("Setting mail: $newEmailAddress")
        Set-ADUser -Identity $sAMAccountName -EmailAddress $newEmailAddress -Server $adServer -Credential $adCredentials
        
        $GroupsToAdd = @("CE-GS-Office365Online", "CE-AS-M365-Base", "CE-GS-RemoteEmployees")
        Add-UserToGroups -ADUser $adUser -Groups $GroupsToAdd

        # Close the exchange session
        if ($exchSession) { Remove-PSSession $exchSession }
    }


    ### Set home drive permissions
    #Set-HomeDrivePermissions


    ### GP accounts can be disabled with no password sent
    if ($employeeType -ine "GP Medtech 32")
    {
        ### Send the initial password
        Send-UserPassword
    }
    else
    {
        Disable-ADAccount -Identity $sAMAccountName
    }

}
catch {
     WriteToLog("Error: Item = $($_.Exception.ItemName) -> Message = $($_.Exception.Message)")
     throw "Error occurred during after account creation logic."
}