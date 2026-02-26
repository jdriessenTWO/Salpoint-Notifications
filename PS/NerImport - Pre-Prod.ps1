# PowerShell script for NERM bulk profile import

# Define API Endpoint and Token
$apiUrl = "https://healthnz.nonemployee.com"
$bearerToken = "ne-OkHpQPLBqUopXoO6Gj2cLwLfLoJQ49IpkYLxRNkyAOhzOYwcub4YwG6Jxrl95Zk1IinH32UKofWQGLu7JyIEdDTKpp5uDu5whuHYlhDE3xDTWheAeeIdegRpsheubbRZ"
$url = "https://healthnz.nonemployee.com/api/profiles?name="

$districtProfileId = "ef6fcd3d-f575-4b3d-bc5e-9ea9853e7e46"
$locationProfileId = "9dc6ee50-40c7-4389-a1b7-7c6f6edbdbb3"
$departmentProfileId = "07e40e76-1dc7-4d76-b345-a1229174edad"
$entityProfileId = "ae2621c4-1dde-4b31-a112-05ff6d3b1435"
$positionTitleProfileId = "72ee7242-9396-45d7-b227-4ed6f0449a62"
$employeeTypeProfileId = "17478bce-927d-45b0-ba15-bd23fcfd1486"
$personProfileId = "e1ea7701-d8e3-4b4b-887e-775f3dcb5712"
$rolesProfileId = "4406799e-9e6b-4c68-a0de-fc1ef5986992"
$organisationProfileId = "45c0e96b-5d1d-4dc7-a89d-1df62a4b74f7"

$districtUrl = "https://healthnz.nonemployee.com/api/profiles?profile_type_id=$districtProfileId&name="

function Get-AllUsers {
    param (
        [string]$bearerToken
    )

    # Base URL for fetching profiles
    $limit = 500  # Adjust limit as needed

    # Fetch all users with metadata
    $userCache = @()
    $userOffset = 0
    $nextUrl = "https://healthnz.nonemployee.com/api/users?limit=$limit&offset=$userOffset&metadata=true"   

    $Headers = @{
        "Authorization" = "Bearer $bearerToken"  # If authentication is required
        "Content-Type"  = "application/json"
        "Accept" = "application/json"
    }

    # Initialize an empty array to store results
    $results = @()

    do {
        # Fetch data from the API
        $userResponse = Invoke-RestMethod -Uri $nextUrl -Method GET -Headers $Headers -ErrorAction Continue
        
        # Append the results to the array
        if ($userResponse.users) {
            foreach ($user in $userResponse.users) {
                $results += [PSCustomObject]@{
                    Name = $user.name
                    ID   = $user.id
                    Email = $user.email
                }
            }
        }
        
       
        # Get the next page URL if available
        if ($results.count -lt $userResponse._metadata.total) {
            $nextUrl = "https://healthnz.nonemployee.com/api$($userResponse._metadata.next)&metadata=true"
            Write-Host "Fetching next page: $nextUrl"
        } else {
            $nextUrl = $null
        }

    } while ($nextUrl)

    return $results
}

function Get-AllProfiles {
    param (
        [string]$ProfileTypeId,
        [string]$bearerToken
    )

    $Headers = @{
        "Authorization" = "Bearer $bearerToken"  # If authentication is required
        "Content-Type"  = "application/json"
        "Accept" = "application/json"
    }

    # Initialize an empty array to store results
    $results = @()
    $baseUrl = "$($apiUrl)/api/profiles"
    $nextUrl = "$($baseUrl)?profile_type_id=$profileTypeId&metadata=true&limit=500"   

    do {
        # Fetch data from the API
        $profileQuery = Invoke-RestMethod -Uri $nextUrl -Method GET -Headers $Headers -ErrorAction Continue
        if($ProfileTypeId -eq $personProfileId) {
            if ($profileQuery.profiles) {
                foreach ($profile in $profileQuery.profiles) {
                    $results += [PSCustomObject]@{
                        Name = $profile.name
                        ID   = $profile.id
                        first_name = $profile.attributes.first_name
                        last_name = $profile.attributes.last_name
                        wpn = $profile.attributes.wpn
                    }
                }
            }
        } else {
             # Append the results to the array
            if ($profileQuery.profiles) {
                foreach ($profile in $profileQuery.profiles) {
                    $results += [PSCustomObject]@{
                        Name = $profile.name
                        ID   = $profile.id
                    }
                }
            }
        }
       
        # Get the next page URL if available
        if ($results.count -lt $profileQuery._metadata.total) {
            $nextUrl = "$apiUrl/api$($profileQuery._metadata.next)&profile_type_id=$ProfileTypeId&metadata=true"
            Write-Host "Fetching next page: $nextUrl"
        } else {
            $nextUrl = $null
        }

    } while ($nextUrl)

    return $results
}

# Function to handle response
function Handle-PostResponse {
    param (
        [string]$postResponse,
        [string]$profileType,
        [string]$wpn,
        [string]$organisationPs,
        [string]$entityPs,
        [string]$personRrAssignment,
        [string]$reportingManagersEmail,
        [string]$requestorsEmail,
        [string]$startDate,
        [string]$districtPs,
        [string]$department,
        [string]$positionTitlePs,
        [string]$primaryWorkingLocationPs,
        [string]$approver_ps,
        [string]$firstName,
        [string]$lastName,
        [int]$rowNumber,
        [System.IO.StreamWriter]$errorWriter
    )
    
    $profileType = $record.profileType
    $outputData = $postResponse  -replace "`r`n", ' '
    $outputData = $outputData  -replace ",", '.'
    if ($postResponse.Contains("Failed: HTTP error code") -or  $postResponse.Contains("failure")) {
        $errorDetails = "Failed to create profile: $wpn. Reason: $outputData"
        Log-Error -wpn $wpn -profileType $profileType -isPrimaryAssignment "Yes" -organisationPs $organisationPs -entityPs $entityPs -personRrAssignment $personRrAssignment -reportingManagersEmail $reportingManagersEmail -requestorsEmail $requestorsEmail -startDate $startDate -endDate $($record.end_date) -districtPs $districtPs -department $department -positionTitlePs $positionTitlePs -primaryWorkingLocationPs $primaryWorkingLocationPs -approver_ps $approver_ps -employeeNo $($record.employee_no) -firstName $firstName -lastName $lastName -type $($record.employment_type) -rowNumber $rowNumber -errorDetails $errorDetails -errorWriter $errorWriter
    } elseif ($postResponse.Contains("Success")) {
        Write-Host "Profile created successfully: $wpn"
    }
}

# Function to log error
function Log-Error {
    param (
        [string]$wpn,
        [string]$profileType,
        [string]$isPrimaryAssignment,
        [string]$organisationPs,
        [string]$entityPs,
        [string]$type,
        [string]$personRrAssignment,
        [string]$reportingManagersEmail,
        [string]$requestorsEmail,
        [string]$startDate,
        [string]$endDate,
        [string]$districtPs,
        [string]$department,
        [string]$positionTitlePs,
        [string]$primaryWorkingLocationPs,
        [string]$approver_ps,
        [string]$employeeNo,
        [string]$firstName,
        [string]$lastName,
        [int]$rowNumber,
        [string]$errorDetails,
        [System.IO.StreamWriter]$errorWriter
    )

    $errorWriter.WriteLine("$wpn,$profileType,$approver_ps,$cost_code,$perm_cost_code,$type,$endDate,$firstName,$isPrimaryAssignment,$lastName,$organisationPs,$entityPs,$personRrAssignment,$reportingManagersEmail,$requestorsEmail,$startDate,$districtPs,$department,$positionTitlePs,$primaryWorkingLocationPs,$employeeNo,$rowNumber,$errorDetails")
    Write-host "$wpn,$profileType,$approver_ps,$cost_code,$perm_cost_code,$type,$endDate,$firstName,$isPrimaryAssignment,$lastName,$organisationPs,$entityPs,$personRrAssignment,$reportingManagersEmail,$requestorsEmail,$startDate,$districtPs,$department,$positionTitlePs,$primaryWorkingLocationPs,$employeeNo,$rowNumber,$errorDetails"
}

function Create-ProfileObject {
    param (
        [string]$profileTypeId,
        [string]$currentTimestamp,
        [string]$profileName,
        [Hashtable]$attributesMap
    )

    # Create a PSCustomObject for the attributes directly from attributesMap
    $attributesObj = [PSCustomObject]@{}

    # Populate attributesObj using the attributesMap
    foreach ($key in $attributesMap.Keys) {
        $attributesObj | Add-Member -MemberType NoteProperty -Name $key -Value $attributesMap[$key]
    }

    # Create the profile object
    $profileObject = [PSCustomObject]@{
        profile_type_id     = $profileTypeId
        id_proofing_status   = "pending"
        status               = "Active"
        name                 = $profileName
        created_at           = $currentTimestamp
        updated_at           = $currentTimestamp
        attributes           = $attributesObj
    }

    # Wrap the profile object in a profiles array and return
    return @{
        profile = $profileObject
    } | ConvertTo-Json -Depth 5
}

function Process-PersonCSV {
    param (
        [string]$csvFilePath,
        [string]$url,
        [string]$bearerToken,
        [string]$errorCsvFilePath
    )

    $currentTimestamp = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffK")
    # Initialize error file for logging
    $errorWriter = New-Object System.IO.StreamWriter($errorCsvFilePath, $true)
    $errorWriter.WriteLine("wpn,profileType,approver_ps,cost_code,perm_cost_code,employment_type,end_date,first_name,is_primary_assignment,last_name,organisation_ps,entity_ps,person_rr_assignment,reporting_managers,requestors_email,start_date,district_ps,department,position_title_ps,primary_working_location_ps,employee_no,row_number,error_details")

    # Read the CSV file
    $csvData = Import-Csv -Path $csvFilePath

    foreach ($record in $csvData) {
        try {
            $wpn = $record.wpn
            $profileType = $record.profileType
            $organisationPs = $record.organisation_ps
            $entityPs = $record.entity_ps
            $personRrAssignment = $record.person_rr_assignment
            $reportingManagersEmail = $record.reporting_managers
            $requestorsEmail = $record.requestors_email
            $startDate = $record.start_date
            $districtPs = $record.district_ps
            $department = $record.department
            $positionTitlePs = $record.position_title_ps
            $primaryWorkingLocationPs = $record.primary_working_location_ps
            $external_org_name = $record.external_org_name
            $approver_ps  = $record.approver_ps
            $rowNumber = $csvData.IndexOf($record) + 1

            # Get Profile Type Id
            $profileTypeId = Get-ProfileTypeId -profileType $profileType

            if (-not $profileTypeId) {
                Log-Error -wpn $wpn -profileType $profileType -isPrimaryAssignment "Yes" -organisationPs $organisationPs -entityPs $entityPs -personRrAssignment $personRrAssignment -reportingManagersEmail $reportingManagersEmail -requestorsEmail $requestorsEmail -startDate $startDate -endDate $($record.end_date) -districtPs $districtPs -department $department -positionTitlePs $positionTitlePs -primaryWorkingLocationPs $primaryWorkingLocationPs -approver_ps $approver_ps -employeeNo $($record.employee_no) -firstName $($record.first_name) -lastName $($record.last_name) -type $($record.employment_type) -rowNumber $rowNumber -errorDetails "Invalid profileType: $profileType" -errorWriter $errorWriter
                continue
            }

            # Create Profile Name
            $profileName = Create-ProfileName -record $record -profileType $profileType -wpn $wpn
            $encodedProfileName = [System.Net.WebUtility]::UrlEncode($profileName)
            if($profileType -eq "person") {
                $currDupe = $persons | Where-Object { ($_.first_name -eq $record.first_name) -and ($_.last_name -eq $record.last_name) -and ($_.wpn -eq $record.wpn) }
            } elseif ($profileType -eq "roles") {
                $currDupe = $roles | Where-Object { ($_.first_name -eq $record.first_name) -and ($_.last_name -eq $record.last_name) -and ($_.wpn -eq $record.wpn) }
            }
            $currDupe
            # Ensure null or empty API response doesn't block execution
            if($currDupe) {
                # Log the error
                $errorDetails = "Duplicate profile found: $($currDupe.ID)"
                Log-Error -wpn $wpn -profileType $profileType -isPrimaryAssignment "Yes" -organisationPs $organisationPs -entityPs $entityPs -personRrAssignment $personRrAssignment -reportingManagersEmail $reportingManagersEmail -requestorsEmail $requestorsEmail -startDate $startDate -endDate $($record.end_date) -districtPs $districtPs -department $department -positionTitlePs $positionTitlePs -primaryWorkingLocationPs $primaryWorkingLocationPs -approver_ps $approver_ps -employeeNo $($record.employee_no)  -firstName $($record.first_name) -lastName $($record.last_name) -type $($record.employment_type) -rowNumber $rowNumber -errorDetails $errorDetails -errorWriter $errorWriter
            
                continue  # Skip to the next record
            }
            
            # If API response is null or empty, continue processing without blocking
            Write-Host "No duplicate profile found or API response was empty. Continuing processing..."            
            # Create attributes map
            $headerMap = @{}
            foreach ($property in $record.PSObject.Properties) {
                $headerMap[$property.Name] = $property.Value
            }
            $attributesMap = Create-AttributesMap -headerMap $headerMap -record $record -bearerToken $bearerToken

            $profileObj = Create-ProfileObject -profileTypeId $profileTypeId -currentTimestamp $currentTimestamp `
                                               -profileName $profileName -attributesMap $attributesMap
# Write-host "2211111122"
            # Send POST request
            $apiUrl = "https://healthnz.nonemployee.com/api/profile"
            $postResponse = Post-Request -url $apiUrl -bearerToken $bearerToken -jsonInputString $profileObj
            $postResponse
# Write-host "2224444444444442"
            # Handle response
            Handle-PostResponse -postResponse $postResponse -wpn $wpn profileType $profileType -organisationPs $organisationPs `
                                 -entityPs $entityPs -personRrAssignment $personRrAssignment `
                                 -reportingManagersEmail $reportingManagersEmail -requestorsEmail $requestorsEmail `
                                 -startDate $startDate -endDate $($record.end_date) -districtPs $districtPs -department $department `
                                 -positionTitlePs $positionTitlePs -primaryWorkingLocationPs $primaryWorkingLocationPs `
                                 -approver_ps $approver_ps -employeeNo $($record.employee_no) -firstName $($record.first_name) -lastName $($record.last_name) -type $($record.employment_type)  -rowNumber $rowNumber -errorWriter $errorWriter
            # Write-host "2222"
        } catch {
            Log-Error -wpn $wpn -profileType $profileType -isPrimaryAssignment "Yes" -organisationPs $organisationPs -entityPs $entityPs -personRrAssignment $personRrAssignment -reportingManagersEmail $reportingManagersEmail -requestorsEmail $requestorsEmail -startDate $startDate -endDate $($record.end_date) -districtPs $districtPs -department $department -positionTitlePs $positionTitlePs -primaryWorkingLocationPs $primaryWorkingLocationPs -approver_ps $approver_ps -employeeNo $($record.employee_no) -firstName $($record.first_name) -lastName $($record.last_name) -type $($record.employment_type) -rowNumber $rowNumber -errorDetails ("Error processing record: $($record) Error: $($_.Message)") -errorWriter $errorWriter
        }
    }

    $errorWriter.Close()
}

# Helper functions for actions like logging errors, checking profile duplicates, creating attributes, etc.
function Get-ProfileTypeId {
    param ($profileType)
    switch ($profileType.ToLower()) {
        "person" { return "e1ea7701-d8e3-4b4b-887e-775f3dcb5712" }
        "roles" { return "4406799e-9e6b-4c68-a0de-fc1ef5986992" }
        default { return $null }
    }
}

function Create-ProfileName {
    param ($record, $profileType, $wpn)
    $firstName = $record.first_name
    $lastName = $record.last_name
    if ($profileType -eq "Roles") {
        return "$firstName $lastName $wpn"
    } else {
        return "$firstName $lastName"
    }
}

function Is-DuplicateProfile {
    param (
        $apiResponse,
        [string]$wpn
    )

    if (-not $apiResponse) {
        Write-Host "API response is empty or null. Duplicate check skipped."
        return $false
    }

    try {
        $responseJson = $apiResponse[0] | ConvertFrom-Json

        # Write-Host "Before first if block"
        # Ensure 'profiles' array exists and has at least one entry
        if ($responseJson.PSObject.Properties.Name -contains "profiles" -and $responseJson.profiles.Count -gt 0) {
            $existingProfile = $responseJson.profiles[0]  # Take the first profile from the array
            
            # Write-Host "first if block executed"
            # Ensure 'attributes' exists and is not null
            if ($existingProfile.PSObject.Properties.Name -contains "attributes" -and $null -ne $existingProfile.attributes) {
                # Write-Host "if block executed"

                # Ensure 'wpn' exists inside 'attributes' and check for a match
                if ($existingProfile.attributes.PSObject.Properties.Name -contains "wpn" -and $existingProfile.attributes.wpn -eq $wpn) {
                    Write-Host "Duplicate found for WPN: $wpn"
                    return $true
                } else {
                    Write-Host "Duplicate not found for WPN: $wpn"
                }
            } else {
                Write-Host "Attributes not found in the first profile entry."
            }
        } else {
            Write-Host "No profiles found in API response."
        }
    } catch {
        Write-Host "Error parsing API response: $_"
    }

    return $false
}

function Create-AttributesMap {
    param (
        [Hashtable]$headerMap,
        [PSCustomObject]$record,
        [string]$bearerToken
    )

    $attributesMap = @{}

    foreach ($header in $headerMap.Keys) {
        $value = $record.$header

        # Process only non-empty values
        if ($value -and $value.Trim() -ne "") {
            if ($header -in @("organisation_ps", "district_ps", "department", "position_title_ps", "primary_working_location_ps", "entity_ps","person_rr_assignment","employment_type")) {

                if($header -eq "primary_working_location_ps") {
                    $foundData = $Locations | Where-Object { $_.Name -eq $value}
                    $attributesMap[$header] = $foundData.id
                }elseif($header -eq "department") {
                    $foundData = $departments | Where-Object { $_.Name -eq $value}
                    $attributesMap[$header] = $foundData.id
                }elseif($header -eq "entity_ps") {
                    $foundData = $Entity | Where-Object { $_.Name -eq $value}
                    $attributesMap[$header] = $foundData.id
                }elseif($header -eq "position_title_ps") {
                    $foundData = $positions | Where-Object { $_.Name -eq $value}
                    $attributesMap[$header] = $foundData.id
                }elseif($header -eq "district_ps") {
                    $foundData = $Districts | Where-Object { $_.Name -eq $value}
                    $attributesMap[$header] = $foundData.id
                }elseif($header -eq "person_rr_assignment") {
                    $foundData = $persons | Where-Object { $_.Name -eq $value -and $_.wpn -eq $record.wpn }
                    $attributesMap[$header] = $foundData.id
                }elseif($header -eq "organisation_ps") {
                    $foundData = $organisations | Where-Object { $_.Name -eq $value}
                    $attributesMap[$header] = $foundData.id
                }elseif($header -eq "employment_type") {
                    $foundData = $employeeType | Where-Object { $_.Name -eq $value}
                    $attributesMap[$header] = $foundData.id 
                }else{
                    $url = "https://healthnz.nonemployee.com/api/profiles?name="
                    try {
                        $responseProfileName = $null
                        $responseStatusCode = 0
    
                        # Attempt 1: Search with uppercase profile name
                        $encodedProfileName = [System.Web.HttpUtility]::UrlEncode($value.ToUpper())
                        $responseProfileName, $responseStatusCode = Get-ApiResponse "$url$encodedProfileName" $bearerToken
    
                        # Process response for Attempt 1
                        if ($responseStatusCode -eq 200 -and $responseProfileName -and $responseProfileName -ne "NotFound" -and $responseProfileName -ne "BadRequest" -and -not $responseProfileName.Contains("Not a valid API route")) {
                            Process-ProfileResponse $responseProfileName $header $attributesMap
                        } else {
                            # Attempt 2: Retry with Camel Case
                            $camelCaseValue = Convert-CamelCase $value
                            $encodedProfileNameCamelCase = [System.Web.HttpUtility]::UrlEncode($camelCaseValue)
                            $responseProfileName, $responseStatusCode = Get-ApiResponse "$url$encodedProfileNameCamelCase" $bearerToken

                            Process-ProfileResponse $responseProfileName $header $attributesMap
    
                            if ($responseStatusCode -ne 200) {
                                # Attempt 3: Retry with original value
                                $encodedProfileNameCsv = [System.Web.HttpUtility]::UrlEncode($value)
                                $responseProfileName, $responseStatusCode = Get-ApiResponse "$url$encodedProfileNameCsv" $bearerToken
    
                                Process-ProfileResponse $responseProfileName $header $attributesMap
                            }
                        }
    
                        # Final check if all attempts failed
                        if ($responseStatusCode -ne 200 -or $responseProfileName -eq "NotFound" -or $responseProfileName -eq "BadRequest" -or $responseProfileName.Contains("Not a valid API route")) {
                            Write-Host "Error: Profile not found for '$value' after all attempts"
                            $attributesMap[$header] = $value
                        }
                    } catch {
                        Write-Host "Error fetching ID for $header with value: '$value' - $_"
                        $attributesMap[$header] = $value
                    }
                }                
            } <# elseif ($header -in @("approver_ps", "reporting_manager")) {       
                
                $foundData = $users | Where-Object { $_.email -eq $value}
                $attributesMap[$header] = $foundData.id
            } #> else {
                $attributesMap[$header] = $value
            }
        }
    }
    if($foundData) {Remove-Variable foundData}
    return $attributesMap
}

function Process-ProfileResponse {
    param (
        [string]$responseProfileName,
        [string]$header,
        [hashtable]$attributesMap
    )

    # Convert JSON string to PowerShell object
    $data = $responseProfileName | ConvertFrom-Json
    $main = $data.profiles
    ### $data.profiles.count -ne 1
    if($data.profiles.count -gt 1) {
        $main = $data.profiles | Where-Object {$_.attributes.wpn -contains $wpn} 
    }
    
    if ($main -and $main.Count -gt 0 -and $main.id) {
        $attributesMap[$header] = $main.id
    } elseif ($data -and $data.users -and $data.users.Count -gt 0 -and $data.users[0].id) {
        $attributesMap[$header] = $data.users[0].id
    } else {
        Write-Host "Warning: ID missing in API response for '$header'"
        $attributesMap[$header] = $header  # You can change this to $value if needed
    }
}



# Function to authenticate using Bearer Token
# Authenticate with Bearer Token
function Authenticate-With-BearerToken {
    param (
        [string]$apiUrl,
        [string]$bearerToken
    )

    try {
        $request = [System.Net.HttpWebRequest]::Create($apiUrl)
        $request.Method = "GET"
        $request.Headers.Add("Authorization", "Bearer $bearerToken")
        $response = $request.GetResponse()

        $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
        $responseContent = $reader.ReadToEnd()
        $reader.Close()
        return $responseContent
    }
    catch {
        Write-Error "Error: $_"
    }
}
# Convert to Camel Case
function Convert-CamelCase{
    param (
        [string]$inputstr
    )
    # Write-Host "val to convert camelcase" $inputstr
    $inputstr = $inputstr.Trim()
    $words = $inputstr -split "\s+"
    $result = ""

    foreach ($word in $words) {
        if ($word -ne "") {
            $result += ($word.Substring(0, 1).ToUpper() + $word.Substring(1).ToLower()) + " "
        }
    }

    if ($result.Length -gt 0) {
        $result = $result.Substring(0, $result.Length - 1)
    }

    return $result
}

# Function to GET API Response
function Get-ApiResponse {
    param (
        [string]$url,
        [string]$bearerToken
    )
    # Write-Host "fetching response for $url with value of token:" $bearerToken
    try {
        # Create HttpClient
        $httpClient = New-Object System.Net.Http.HttpClient
        $httpClient.DefaultRequestHeaders.Authorization = "Bearer $bearerToken"
        $httpClient.DefaultRequestHeaders.Accept.Add("application/json")

        # Send GET request
        $response = $httpClient.GetAsync($url).Result
        Write-Host "Fetching response for $url : Status Code - $($response.StatusCode)"

        # Check response status
        if ($response.StatusCode -eq 200) {
            # Write-Host "Status Code - 200"
            $responseProfileName = $response.Content.ReadAsStringAsync().Result
            return $responseProfileName, $response.StatusCode  # Return both values
        } elseif ($response.StatusCode -eq 400)  {
            $responseProfileName = $response.Content.ReadAsStringAsync().Result
            return $responseProfileName, $response.StatusCode  # Return both values
        } elseif ($response.StatusCode -eq "NotFound") {
            Write-Host "Status Code - Not Found"
            $responseProfileName = $response.Content.ReadAsStringAsync().Result
            return $responseProfileName, $response.StatusCode  # Return both values
        } else {
            throw "Failed: HTTP error code : $($response.StatusCode)"
        }
    }
    catch {
        Write-Error "Error fetching API response: $_"
        return $null, 500  # Return null and status code 500 in case of error
    }
    finally {
        # Dispose the HttpClient
        if ($httpClient -ne $null) {
            $httpClient.Dispose()
        }
    }
}

# Function to PATCH Request
# Send PATCH Request
function Patch-Request {
    param (
        [string]$urlString,      # URL to which the PATCH request will be sent
        [string]$bearerToken,      # Bearer token for authorization
        [string]$jsonInputString  # JSON payload for the PATCH request

    )

    try {
        # Set the headers for the PATCH request
        $headers = @{
            "Authorization" = "Bearer $bearerToken"
            "Content-Type"  = "application/json"
            "Accept"        = "*/*"
        }
        # Write-host "Pause"
        # Send the PATCH request using Invoke-RestMethod
        $response = Invoke-RestMethod -Uri $urlString -Method Patch -Headers $headers -Body $jsonInputString -ContentType "application/json" -StatusCodeVariable responseStatusCode

        # Print the status code to verify success (response should be an object in PowerShell)

        if ($responseStatusCode -eq 200) {
            return $response
        } else {
            throw "Failed: HTTP error code : $responseStatusCode"
        }

    } catch {
        Write-Host "Error occurred: $_"
    }

    return "success"
}

# Function to Process CSV for Profile Update (Matching Java Logic)
function Process-CsvFile {
    param (
        [string]$csvFilePath,
        [string]$url,
        [string]$locationProfileId,
        [string]$departmentProfileId,
        [string]$entityProfileId,
        [string]$positionTitleProfileId,
        [string]$bearerToken,
        [string]$errorCsvFilePath
    )
    
    # Create error CSV file header if it does not exist
    if (-not (Test-Path $errorCsvFilePath)) {
        "ProfileName,AttributeName,ErrorDetails" | Out-File -FilePath $errorCsvFilePath -Encoding UTF8
    }

    $totalRowCount = 0
    $valueIdCounter = 0
    $mapValueList = @{}
    $listLocations = @()
    $listDepartments = @()
    $listEntity = @()
    $listPositionTitle = @()
    $listEmployeeType = @()
    $profileId = $null
    $attributeName = ""

    # Read the CSV file
    $csvContent = Import-Csv -Path $csvFilePath

    foreach ($record in $csvContent) {
        Write-Host "**************record***********" $record
        $totalRowCount++  # Increment the row count
        $profileName = $record.ProfileName
        $attributeName = $record.AttributeName
        $value = $record.Value
       
        $resultMapProfileName = $Districts | Where-Object { $_.Name -eq $profileName}
        $profileId = $resultMapProfileName.id

        if($attributeName -eq "permitted_locations") {
            $foundData = $Locations | Where-Object { $_.Name -eq $value}
        }elseif($attributeName -eq "permitted_department") {
            $foundData = $departments | Where-Object { $_.Name -eq $value}
        }elseif($attributeName -eq "permitted_entity") {
            $foundData = $Entity | Where-Object { $_.Name -eq $value}
        }elseif($attributeName -eq "permitted_position_title") {
            $foundData = $positions | Where-Object { $_.Name -eq $value}
        }elseif($attributeName -eq "permitted_employment_type") {
            $foundData = $employeeType | Where-Object { $_.Name -eq $value}
        }else{
            Write-host "Wrong type please try again"
        }

        if(($attributeName -eq "permitted_locations") -and ($listLocations -notcontains $foundData.id)) {
            $listLocations += $foundData.id
        }elseif(($attributeName -eq "permitted_department") -and ($listDepartments -notcontains $foundData.id)) {
            $listDepartments += $foundData.id
        }elseif(($attributeName -eq "permitted_entity") -and ($listEntity -notcontains $foundData.id)) {
            $listEntity += $foundData.id
        }elseif(($attributeName -eq "permitted_position_title") -and ($listPositionTitle -notcontains $foundData.id)) {
            $listPositionTitle += $foundData.id
        }elseif(($attributeName -eq "permitted_employment_type") -and ($listEmployeeType -notcontains $foundData.id)) {
            $listEmployeeType += $foundData.id
        }
        $valueIdCounter++

    }

    # Create payload
    if($listLocations.Length -gt 0) {
        $mapValueList["permitted_locations"] += $listLocations -join ", "
    }
    if($listDepartments.Length -gt 0) {
        $mapValueList["permitted_department"] += $listDepartments -join ", "
    }
    if($listEntity.Length -gt 0) {
        $mapValueList["permitted_entity"] += $listEntity -join ", "
    }
    if($listPositionTitle.Length -gt 0) {
        $mapValueList["permitted_position_title"] += $listPositionTitle -join ", "
    }
    if($listEmployeeType.Length -gt 0) {
        $mapValueList["permitted_employment_type"] += $listEmployeeType -join ", "
    }
    
    Write-Host "Total CSV rows processed: $totalRowCount"

    $attributeObj = @{
        attributes = $mapValueList
    }

    $jsonPayload = @{
        profile = $attributeObj
    } | ConvertTo-Json 

    Write-Host "jsonPayload: $jsonPayload"
    $url = "https://healthnz.nonemployee.com/api/profiles/$profileId"
    # Call patch

    Patch-Request "$url" $bearerToken $jsonPayload
}

# Function to Process CSV for Sub-Organisation Update (Matching Java Logic)
function Process-SubOrgCsvFile {
    param ([string]$csvPath)
    $data = Import-Csv -Path $csvPath
    foreach ($row in $data) {
        Write-Host "Processing sub-org update for profile: $($row.profileName)"
        
        $profileName = [System.Web.HttpUtility]::UrlEncode($row.profileName.ToUpper())
        $response = Get-ApiResponse "$apiUrl$profileName"
        
        if ($null -eq $response -or $response.profiles.Count -eq 0) {
            Write-Host "Profile not found for: $profileName"
            continue
        }
        
        $profileId = $response.profiles[0].id
        
        if ($null -ne $response.profiles[0].attributes."sub-organisation") {
            $subOrgValue = $response.profiles[0].attributes."sub-organisation"
            
            $encodedSubOrg = [System.Web.HttpUtility]::UrlEncode($subOrgValue)
            $subOrgResponse = Get-ApiResponse "$apiUrl$encodedSubOrg"
            
            if ($null -eq $subOrgResponse -or $subOrgResponse.profiles.Count -eq 0) {
                Write-Host "Sub-organisation not found: $subOrgValue"
                continue
            }
            
            $subOrgId = $subOrgResponse.profiles[0].id
            
            $patchData = @{ profile = @{ attributes = @{ "entity_ps" = $subOrgId } } } | ConvertTo-Json -Depth 10
            
            Write-Host "Updating profile ID: $profileId with entity_ps ID $subOrgId"
            Patch-Request -url "https://healthnz.nonemployee.com/api/profiles/$profileId" -jsonInputString $patchData
        }
    }
}
function Post-Request {
    param (
        [string]$urlString,
        [string]$bearerToken,
        $jsonInputString
    )

    try {
        # Prepare headers
        $headers = @{
            "Authorization" = "Bearer $bearerToken"
            "Content-Type"  = "application/json"
            "Accept"        = "*/*"
        }

        # Perform the POST request
        # Write-Host "Sending POST request to: $urlString"
        $response = Invoke-RestMethod -Uri $urlString -Headers $headers -Method Post -Body $jsonInputString -ContentType "application/json" -StatusCodeVariable responseStatusCode

        # Write-Host "Response status code: $responseStatusCode"

        # Check if response status code is 200 (OK) or 201 (Created)
        if ($responseStatusCode -eq 200 -or $responseStatusCode -eq 201) {
            return $response
        } else {
            throw "Failed: HTTP error code: $responseStatusCode"
        }
    } catch {
        Write-Host "Error occurred: $_"
        return "failure record not found: $_"
    }
}


# Main logic

Write-Host "********url******* $url"
Write-Host "********districtUrl******* $districtUrl"

try {
    Write-Host "Select Object Type"
    Write-Host "1. NERM Person Records"
    Write-Host "2. NERM Role Records"
    Write-Host "3. Nerm Permitted Values"
    
    $input = Read-Host "Enter the number of your choice (1-3)"

    $response = Authenticate-With-BearerToken -apiUrl $apiUrl -bearerToken $bearerToken

    switch ($input)
    {
        1 {
        <#
            Add a Person's Profile to Nerm
        #> 
    
        Write-Host "Selected upload NERM Person records"
        $global:persons = Get-AllProfiles -ProfileTypeId $personProfileId -bearerToken $bearerToken
        $csvFilePath = "C:\Users\kanwaldhanoa\Downloads\South Canterbury - Data Load\South Canterbury - initial load\South Canterbury - person - initial.csv"
        $errorcsvFilePath = "C:\Users\kanwaldhanoa\Downloads\South Canterbury - Data Load\South Canterbury - initial load\Error files\person-error.csv"
        
        Process-PersonCSV -csvFilePath $csvFilePath -url $url -bearerToken $bearerToken -errorcsvFilePath $errorcsvFilePath
        }
        2 {
        <#
            Add Roles to a Person Profile
        #> 
        
        Write-Host "Selected upload NERM Roles records"

        Write-Host "Starting load of all Data"
        Write-Host "positions"
        $global:positions = Get-AllProfiles -ProfileTypeId $positionTitleProfileId -bearerToken $bearerToken
        Write-Host "departments"
        $global:departments = Get-AllProfiles -ProfileTypeId $departmentProfileId -bearerToken $bearerToken
        Write-Host "Locations"
        $global:Locations = Get-AllProfiles -ProfileTypeId $locationProfileId -bearerToken $bearerToken
        Write-Host "Districts"
        $global:Districts = Get-AllProfiles -ProfileTypeId $districtProfileId -bearerToken $bearerToken
        Write-Host "Entity"
        $global:Entity = Get-AllProfiles -ProfileTypeId $entityProfileId -bearerToken $bearerToken
        Write-Host "persons"
        $global:persons =Get-AllProfiles -ProfileTypeId $personProfileId -bearerToken $bearerToken
        Write-Host "roles"
        $global:roles = Get-AllProfiles -ProfileTypeId $rolesProfileId -bearerToken $bearerToken
        Write-Host "organisation"
        $global:organisations = Get-AllProfiles -ProfileTypeId $organisationProfileId -bearerToken $bearerToken
        Write-Host "employment_type"
        $global:employeeType = Get-AllProfiles -ProfileTypeId $employeeTypeProfileId -bearerToken $bearerToken
        <# Write-Host "Users"
        $global:users = Get-AllUsers -bearerToken $bearerToken #>
        Write-Host "Done"

        $csvFilePath = "C:\Users\kanwaldhanoa\Downloads\South Canterbury - Data Load\South Canterbury - initial load\Error files\roles2-error.csv"
        $errorcsvFilePath = "C:\Users\kanwaldhanoa\Downloads\South Canterbury - Data Load\South Canterbury - initial load\Error files\roles3-error.csv"

        Process-PersonCSV -csvFilePath $csvFilePath -url $url -bearerToken $bearerToken -errorcsvFilePath $errorcsvFilePath
        }
        3 {
        <#
            To add permitted values to a district, Update the below csvFilePath for each permitted Values
        #> 
        Write-Host "Select Object Type"
        Write-Host "1. Departments"
        Write-Host "2. Locations"
        Write-Host "3. Position Titles"
        Write-Host "4. Entity"
        Write-Host "5. Employment_type"
        
        $Choice = Read-Host "Enter the number of your choice (1-5)"

        switch ($Choice)
    {
        1 {
            Write-Host "departments"
            $global:departments = Get-AllProfiles -ProfileTypeId $departmentProfileId -bearerToken $bearerToken
            Write-Host "Districts"
            $global:Districts = Get-AllProfiles -ProfileTypeId $districtProfileId -bearerToken $bearerToken
        }
        
        2{
            Write-Host "Locations"
            $global:Locations = Get-AllProfiles -ProfileTypeId $locationProfileId -bearerToken $bearerToken
            Write-Host "Districts"
            $global:Districts = Get-AllProfiles -ProfileTypeId $districtProfileId -bearerToken $bearerToken
        }
        3{
            Write-Host "positions"
            $global:positions = Get-AllProfiles -ProfileTypeId $positionTitleProfileId -bearerToken $bearerToken
            Write-Host "Districts"
            $global:Districts = Get-AllProfiles -ProfileTypeId $districtProfileId -bearerToken $bearerToken
        }
        4{
            Write-Host "Entity"
            $global:Entity = Get-AllProfiles -ProfileTypeId $entityProfileId -bearerToken $bearerToken
            Write-Host "Districts"
            $global:Districts = Get-AllProfiles -ProfileTypeId $districtProfileId -bearerToken $bearerToken
        }
        5{
            Write-Host "employment_type"
            $global:employeeType = Get-AllProfiles -ProfileTypeId $employeeTypeProfileId -bearerToken $bearerToken
            Write-Host "Districts"
            $global:Districts = Get-AllProfiles -ProfileTypeId $districtProfileId -bearerToken $bearerToken
        }
    }

        Write-Host "Selected upload Nerm Permitted Values"
        $csvFilePath = "C:\Users\kanwaldhanoa\Downloads\NelsonMarlborough\Nelson Marlborough\Nelson Marlborough - position title - permitted values - update take 2.csv"
        $errorcsvFilePath = "C:\Users\kanwaldhanoa\Downloads\NelsonMarlborough\Nelson Marlborough\Error\position title-permitted-error.csv"
        
        Process-CsvFile -csvFilePath $csvFilePath -url $districtUrl -locationProfileId $locationProfileId -departmentProfileId $departmentProfileId -entityProfileId $entityProfileId -positionTitleProfileId $positionTitleProfileId -bearerToken $bearerToken -errorcsvFilePath $errorcsvFilePath
        }
        Default {"Invalid entry, try again"}
    }

} catch {
    Write-Host "Error: $($_.Exception.Message)"
}
