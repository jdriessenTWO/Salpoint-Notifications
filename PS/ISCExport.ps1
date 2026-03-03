############################################################################
#ISC Person/Roles Extract
############################################################################
#
#
#
############################################################################ 

cls
Import-Module ImportExcel

# Define the API endpoint and headers
$baseUrl = "https://healthnz.api.identitynow.com"
$filePath = "C:\scripts\Waikato"

# Define the token endpoint URL
$tokenEndpoint = "$($baseUrl)/oauth/token"

# Define your client credentials
$clientId = "a5703ecb5ea74350b08d9b18bf7140d8"
$clientSecret = "42b68d41ae57ca8da59b77e0a2a7887a2ef538715dcc9587b40a8899ced9fbdf"

# Construct the URL with client credentials as parameters
$urlWithParams = "$($tokenEndpoint)?grant_type=client_credentials&client_id=$clientID&client_secret=$clientSecret"

# Make the POST request to obtain the access token
$responseToken = Invoke-RestMethod -Uri $urlWithParams -Method Post

# Check if the request was successful and if the access token was obtained
if ($responseToken -and $responseToken.access_token) {
    # Extract the access token from the response
    $accessToken = $responseToken.access_token

    # Use the access token for subsequent API requests
    # Your API requests here
    Write-Host "Access token obtained: $accessToken"
} else {
    Write-Host "Failed to obtain access token."
}
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Bearer $accessToken")
$headers.Add("Content-Type", "application/json")

$source = 'identityProfile.name:"Waikato"'

# Define the search payload (adjust as needed)
 $searchBody = @{
     "query" = @{
         "query"= $source
     }
     "sort" = @("id")
     "indices" = @("identities")
 } | ConvertTo-Json

$limit = 1000
$response = Invoke-RestMethod -Uri "$baseUrl/v3/search?limit=$limit" -Method Post -Headers $headers -Body $searchBody -ResponseHeadersVariable responseHeaders
$identitiesArray = @()
$roleArray = @()
$personArray =@()

$response | ForEach-Object {
    $identitiesArray += $_
}

$lastId = $response[-1].id

# Loop through pages until all data is retrieved
do {
    
    $uri = "$baseUrl/v3/search?limit=$limit"
    $searchBody = @{
        "query" = @{
            "query"= $source
        }
        "sort" = @("id")
        "indices" = @("identities")
        "searchAfter" = @("$lastId")
    } | ConvertTo-Json

    $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $searchBody

    # Add the new records to the accumulated list
    $response | ForEach-Object {
        $identitiesArray += $_
    }

    $lastId = $response[-1].id
    Write-Host $lastId
    Write-Host $identitiesArray.Count

} while ($response.Count -gt 0)

Write-Host "Total identities fetched: $($identitiesArray.Count)"

    # $data = $response.content | ConvertFrom-Json
    $identitiesArray | ForEach-Object {
        Write-Output "Working on name: $($_.name)"
        $personObject = New-Object -TypeName PSObject
        $personObject | Add-Member -MemberType NoteProperty -Name 'wpn' -value $_.attributes.wpnNumber
        $personObject | Add-Member -MemberType NoteProperty -Name 'profileType' -value "Person"
        $personObject | Add-Member -MemberType NoteProperty -Name 'first_name' -value $_.attributes.firstname
        $personObject | Add-Member -MemberType NoteProperty -Name 'last_name' -value $_.attributes.lastname
        $personObject | Add-Member -MemberType NoteProperty -Name 'preferred_name' -value $_.attributes.preferredName
        $personObject | Add-Member -MemberType NoteProperty -Name 'area_of_work' -value ""
        $personObject | Add-Member -MemberType NoteProperty -Name 'common_person_number' -value ""
        $personObject | Add-Member -MemberType NoteProperty -Name 'medical_council_newzealand_number' -value ""
        #$personObject | Add-Member -MemberType NoteProperty -Name 'external_organisation_email' -value $_.attributes.email #Only needed for external people, may be external_organisation_email
        $personArray += $personObject
        
        $roleObject = New-Object -TypeName PSObject
        $roleObject | Add-Member -MemberType NoteProperty -Name 'ISC-LifecycleState' -value $_.attributes.cloudLifecycleState
        $roleObject | Add-Member -MemberType NoteProperty -Name 'ISC-ADDN' -value $_.attributes.adDn 
        $roleObject | Add-Member -MemberType NoteProperty -Name 'wpn' -value $_.attributes.wpnNumber # Required
        $roleObject | Add-Member -MemberType NoteProperty -Name 'profileType' -value "Roles" # Required
        $roleObject | Add-Member -MemberType NoteProperty -Name 'approver_ps' -value "" # Required
        $roleObject | Add-Member -MemberType NoteProperty -Name 'cost_code' -value $_.attributes.costCenter
        $roleObject | Add-Member -MemberType NoteProperty -Name 'perm_cost_code' -value $_.attributes.costCenter
        $roleObject | Add-Member -MemberType NoteProperty -Name 'employment_type' -value $_.attributes.employeeType # Required "type" is the old header and "employment_type" is the new header
        # Write-host "enddate: " $_.attributes.endDate
        if(($_.attributes.endDate) -and ($_.attributes.endDate -ne "never")){
            
            $endDate=[Datetime]::ParseExact($_.attributes.endDate, 'yyyy-MM-dd', $null)
            $endDateStr = $endDate.toString("dd/MM/yyyy")
            $roleObject | Add-Member -MemberType NoteProperty -Name 'end_date' -value $endDateStr
        } else {
            $roleObject | Add-Member -MemberType NoteProperty -Name 'end_date' -value ""
        }
        
        $roleObject | Add-Member -MemberType NoteProperty -Name 'first_name' -value $_.attributes.firstname # Required
        $roleObject | Add-Member -MemberType NoteProperty -Name 'is_primary_assignment' -value "Yes" # Required
        $roleObject | Add-Member -MemberType NoteProperty -Name 'last_name' -value $_.attributes.lastname # Required
        $roleObject | Add-Member -MemberType NoteProperty -Name 'organisation_ps' -value "Health NZ - Te Whatu Ora" # Required
        $roleObject | Add-Member -MemberType NoteProperty -Name 'entity_ps' -value $_.attributes.entity 
        $roleObject | Add-Member -MemberType NoteProperty -Name 'person_rr_assignment' -value "$($_.attributes.firstname) $($_.attributes.lastname)" # Required, populdated directly from WPN
        if($null -ne $_.manager){try {
            Start-Sleep -Milliseconds 50
            $headers.Add("X-SailPoint-Experimental", "true")
            $headers.Add("Accept", "application/json")
            $uri = $baseUrl + '/v2024/identities?filters=id eq "' + $_.manager.id + '"'
            $managerCheck = Invoke-WebRequest -Uri $uri -Method 'GET' -Headers $headers
            $managerData = $managerCheck.content | ConvertFrom-Json
            if($null -ne $managerData) {
                # clear Archived/Terminated.
                $uri = $baseUrl + '/v2024/accounts?count=true&limit=10&offset=0&filters=identityId eq "'+$managerData.id+'"' 
                $managerCheckAccounts = Invoke-WebRequest -Uri $uri -Method 'GET' -Headers $headers
                $managerDataAccounts = $managerCheckAccounts.content | ConvertFrom-Json
                $nermUser = $managerDataAccounts | Where-Object { $_.sourceName -eq "NERM - USER"}
                $roleObject | Add-Member -MemberType NoteProperty -Name 'reporting_manager' -value $nermUser.nativeIdentity
                Clear-Variable -Name managerData
            }
            Clear-Variable -Name managerCheck
            $headers.Remove("X-SailPoint-Experimental")
            $headers.Remove("Accept")
        }
        catch {
            Write-Host "Error occurred: $_"
        }
            
        } else {
            $roleObject | Add-Member -MemberType NoteProperty -Name 'reporting_manager' -value "" # Nice to have 
        }
        $roleObject | Add-Member -MemberType NoteProperty -Name 'requestors_email' -value "identity.notification@tewhatuora.govt.nz" # Static to - svathi.babu@tewhatuora.govt.nz 
        if($_.attributes.startDate){
            $startDate=[Datetime]::ParseExact($_.attributes.startDate, 'yyyy-MM-dd', $null)
            $startDateStr = $startDate.toString("dd/MM/yyyy")
            $roleObject | Add-Member -MemberType NoteProperty -Name 'start_date' -value $startDateStr
        } else {
            $roleObject | Add-Member -MemberType NoteProperty -Name 'start_date' -value ""
        }
        $roleObject | Add-Member -MemberType NoteProperty -Name 'district_ps' -value $_.attributes.district # Required - Match to dynamic list data
        #$roleObject | Add-Member -MemberType NoteProperty -Name 'Preferred_Name' -value $_.attributes.preferredName # Optional, if populated
        $roleObject | Add-Member -MemberType NoteProperty -Name 'department' -value $_.attributes.department #Required - Match to dynamic list data
        $roleObject | Add-Member -MemberType NoteProperty -Name 'position_title_ps' -value $_.attributes.jobTitle #Required - Match to dynamic list data
        $roleObject | Add-Member -MemberType NoteProperty -Name 'primary_working_location_ps' -value $_.attributes.location #Required - Match to dynamic list data
        $roleObject | Add-Member -MemberType NoteProperty -Name 'employee_no' -value $_.attributes.identificationNumber
        $roleObject | Add-Member -MemberType NoteProperty -Name 'external_org_name ' -value ""
        $roleObject | Add-Member -MemberType NoteProperty -Name 'preferred_name' -value $_.attributes.preferredName
        $roleArray += $roleObject

        $personArray.Count

    }

    $personPath = "$filePath\Waikato-person.csv"
    $rolesPath = "$filePath\Waikato-roles.csv"
    # Out-File -InputObject $roleArrayTest -FilePath $rolesPath
    $personArray | Export-Csv -Path $personPath -NoTypeInformation
    $roleArray | Export-Csv -Path $rolesPath -NoTypeInformation