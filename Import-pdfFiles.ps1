<#
.SYNOPSIS
    Imports PDF files and their metadata into Paperless-ngx using its API.

.DESCRIPTION
    This script reads a CSV file containing metadata and uploads corresponding PDF files to a Paperless-ngx instance.
    After upload, it updates custom fields for each document using the API.

.PARAMETER pdfFolder
    The folder containing the PDF files to import.

.PARAMETER csvFile
    The path to the CSV file containing metadata for each PDF.

.PARAMETER paperlessInstance
    The base URL of the Paperless-ngx instance. For example: http://<IP_ADDRESS>:<PORT>

.NOTES
    The csv file should have the following columns, which should be separated by semicolons:
    - InvoiceID
    - InvoiceDate
    - Company
    - Amount
    - Notes

    This represents the custom fields in Paperless-ngx which would be numbered 1-5 in the JSON body (and also in Paperless-ngx ;) ).
    The script assumes that the PDF files are named according to the InvoiceID in the CSV file.
    The script uses Basic Authentication for the Paperless-ngx API.

#>

param (
    [string]$pdfFolder = "C:\TodaysImport\pdf",
    [string]$csvFile = "C:\TodaysImport\Import.csv",
    [string]$paperlessInstance = "http://<IP_ADDRESS>:<PORT>"
)

# Prompt for credentials interactively and store as plain text (for basic auth)
# Note: Storing credentials in plain text is not secure!
$cred = Get-Credential -Message "Enter Paperless-ngx API credentials"
$username = $cred.UserName
$password = $cred.GetNetworkCredential().Password

# Remove trailing slash from $paperlessInstance if present
if ($paperlessInstance.EndsWith("/")) {
    $paperlessInstance = $paperlessInstance.TrimEnd("/")
}

# Paperless-ngx API data
$apiUrl = $paperlessInstance + "/api/documents/post_document/"  # URL of the Paperless-ngx API (adjust if needed)
$StatusURL = $paperlessInstance + "/api/tasks/?task_id="
$UpdateURL = $paperlessInstance + "/api/documents/bulk_edit/"


# Read CSV file
$csvData = Import-Csv -Path $csvFile -Delimiter ";"

# Create Basic Auth header for the API request
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($username):$($password)"))

# Upload each PDF and send classification data
foreach ($row in $csvData) {
    Write-Output "Uploading $($row.InvoiceID)"
    # Find the PDF file that belongs to this entry
    $pdfFile = Join-Path -Path $pdfFolder -ChildPath "$($row.InvoiceID).pdf"
    
    if (Test-Path $pdfFile) {
        # Prepare the multipart form data request
        $formData = @{
            document = Get-Item -Path $pdfFile
            title = $row.InvoiceID
            created = ([datetime]::ParseExact($($row.InvoiceDate), "dd.MM.yyyy", $null)).ToString("yyyy-MM-ddTHH:mm")
        }

        # Send the POST request with the form data
        try {
            $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers @{Authorization = "Basic $base64AuthInfo"} -Form $formdata -ErrorAction Stop
            Start-Sleep -Seconds 1
        }
        catch {
            # Error handling if the request fails
            Write-Host "Error uploading the document for the request: $_"
        }
        
        $CurrentStatusURL = $StatusURL + $response

        # Maximum number of retries
        $maxRetries = 100
        $retryCount = 0

        # Loop to repeatedly check the status, cause it can take a while until the document is processed and the status is "SUCCESS"
        while ($retryCount -lt $maxRetries) {
            try {
                # Send GET request and check the status
                $Status = Invoke-RestMethod -Uri $CurrentStatusURL -Method Get -Headers @{Authorization = "Basic $base64AuthInfo"} -ErrorAction Stop
                
                # If the status is "SUCCESS", exit the loop
                if ($Status.status -eq "SUCCESS") {
                    Write-Host "Status is now: $($Status.status)"
                    break
                }

                # If the status is still "pending", wait 2 seconds and try again
                Start-Sleep -Seconds 2
                $retryCount++
            }
            catch {
                Write-Host "Status of the document is still pending - further processing does not make sense or retrieving the status failed: $_"
                break
            }
        }


        $CurrentID = $Status.related_document
        


        $formDataUpdate = @"
{
    "documents": [
        $CurrentID
    ],
    "method": "modify_custom_fields",
    "parameters":{
        "add_custom_fields": {
"@
        if ($($row.Company) -notlike "") {
            $formDataUpdate += @"

            "1":"$($row.Company)",
"@
        }

        if ($($row.Amount) -notlike "") {
            $formDataUpdate += @"

            "2":"$(($row.Amount -replace '\.', '') -replace ',', '.' )",
"@
        }

        if ($($row.InvoiceDate) -notlike "") {
            $formDataUpdate += @"

            "3":"$(([datetime]::ParseExact($($row.InvoiceDate), "dd.MM.yyyy", $null)).ToString("yyyy-MM-dd"))",
"@
        }

        if ($($row.InvoiceID) -notlike "") {
            $formDataUpdate += @"

            "4":"$($row.InvoiceID)",
"@
        }

        if ($($row.Note) -notlike "") {
            $formDataUpdate += @"

            "5":"$($row.Note)"
"@
        }

        $formDataUpdate += @"

        },
        "remove_custom_fields": {}
    }
}
"@

        $Updateheaders = @{
            "Authorization" = "Basic $base64AuthInfo"
            "Content-Type"  = "application/json"
        }

        try {
            $response = Invoke-RestMethod -Uri $UpdateURL -Method Post -Headers $Updateheaders -Body $formDataUpdate -ErrorAction Stop
        }
        catch {
            # Error handling if the request fails
            Write-Output ""
            Write-Output "Current JSON"
            Write-Output "$formDataUpdate"
            Write-Error "Error updating the custom fields: $_" -ErrorAction Continue
        }
    } else {
        Write-Host "PDF file for $($row.FileName) not found!"
    }
    
}