# Define the base directory for user signatures
$signaturesDirectoryBase = "C:\users"

# Log file setup with dynamic naming to include date and time
$LogFileName = "Log_" + $(Get-Date -Format "yyyyMMddHHmmss") + ".log"
$LogFilePath = Join-Path -Path "c:\users" -ChildPath $LogFileName

Function LogWrite {
    Param ([string]$logstring)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $logstring"
    # Append the log entry to the log file
    $logEntry | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
}

# Function to process each user
Function ProcessUser {
    Param ($user)
    # Specify user-specific directory for signatures
    $signaturesDirectory = Join-Path -Path $signaturesDirectoryBase -ChildPath "$($user.SamAccountName)\AppData\Roaming\Microsoft\Signatures"
    
    # Ensure the directory exists
    if (-not (Test-Path -Path $signaturesDirectory)) {
        New-Item -ItemType Directory -Path $signaturesDirectory
    }

    # Define the basic HTML template for the signature, including the company logo image
    $template = @"
<html>
<body>
    <img src="file://dc10/c$/windows/sysvol_dfsr/sysvol/tgioa.com/scripts/signatures/tgi_files/image001.jpg" alt="Company Logo" style="width: 336px; height: 192px;">
    <p>Best regards,</p>
    <p><strong>{DisplayName}</strong><br>
    {Title}<br>
    Office: {Office}<br>
    Phone: {Phone}<br>
    Address: {PhysicalDeliveryOfficeName}, {StreetAddress}, {City}, {State}, {PostalCode}, {Country}</p>
</body>
</html>
"@

    # Replace placeholders in the template with actual user details
    $signatureContent = $template -replace "\{DisplayName\}", $user.DisplayName `
                                    -replace "\{Title\}", $user.Title `
                                    -replace "\{Office\}", $user.OfficePhone `
                                    -replace "\{Phone\}", $user.OfficePhone `
                                    -replace "\{PhysicalDeliveryOfficeName\}", $user.PhysicalDeliveryOfficeName `
                                    -replace "\{StreetAddress\}", $user.StreetAddress `
                                    -replace "\{City\}", $user.City `
                                    -replace "\{State\}", $user.State `
                                    -replace "\{PostalCode\}", $user.PostalCode `
                                    -replace "\{Country\}", $user.Country

    # Define the file path for the user's signature file
    $filePath = Join-Path -Path $signaturesDirectory -ChildPath "signature.htm"
    # Save the signature content to the file
    $signatureContent | Out-File -FilePath $filePath -Encoding UTF8
    LogWrite "Signature generated for $($user.SamAccountName)"
}

# Fetch users from Active Directory
$users = Get-ADUser -Filter * -Property DisplayName, Title, OfficePhone, PhysicalDeliveryOfficeName, StreetAddress, City, State, PostalCode, Country 

# Process each user
foreach ($user in $users) {
    ProcessUser -user $user
}

LogWrite "All signatures have been generated and configured."
