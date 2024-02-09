

# Check if the Az module is available
$azModule = Get-Module -Name Az -ListAvailable

if ($azModule -eq $null) {
    # The Az module is not installed, proceed with installation
    Write-Host "Azure PowerShell module (Az) is not installed. Installing now..." -ForegroundColor Green
    
    # Install the Az module from the PowerShell Gallery
    Install-Module -Name Az -Scope CurrentUser -AllowClobber -Force

    Write-Host "Azure PowerShell module (Az) installed successfully. Please wait.." -ForegroundColor Yellow -Verbose
} else {
    # The Az module is already installed
    Write-Host "Azure PowerShell module (Az) is already installed." -ForegroundColor Red
}

Import-Module ActiveDirectory

# Function to create an email signature
function Create-EmailSignature {
    param (
        [string]$FolderPath,
        [hashtable]$UserProperties,
        [string]$LogoPath
    )

    if ($null -eq $UserProperties -or $UserProperties.Count -eq 0) {
        Write-Host 'Error: $UserProperties is null or empty. Please ensure it is populated with data from Active Directory.'
        return
    }

    # Debug: Display the properties to be used in the signature
    Write-Host "Preparing to create signature with the following details:"
    $UserProperties.GetEnumerator() | ForEach-Object { Write-Host "$($_.Key): $($_.Value)" }

    # Construct the HTML signature with user properties
    $signatureHTML = @"
    <html>
    <body>
        <p><strong>$($UserProperties['DisplayName'])</strong><br />
        $($UserProperties['Title'])<br />
        Department: $($UserProperties['Department'])<br />
        Company: $($UserProperties['Company'])<br />
        Office: $($UserProperties['Office'])<br />
        Address: $($UserProperties['StreetAddress']), $($UserProperties['City']), $($UserProperties['State']) $($UserProperties['PostalCode'])<br />
        Email: <a href='mailto:$($UserProperties['EmailAddress'])'>$($UserProperties['EmailAddress'])</a></p>
        <img src='$LogoPath' alt='Company Logo' />
    </body>
    </html>
"@

    # Save the signature HTML to a file
    try {
        $signatureHTML | Out-File -FilePath "$FolderPath\Signature.htm" -Encoding UTF8
        Write-Host "Signature created successfully for $($UserProperties['DisplayName'])."
    } catch {
        Write-Host "Failed to save signature: $($_.Exception.Message)"
    }
}

# Fetch the current logged-in user's username
$currentUser = $env:USERNAME

# Perform the AD query to get user details
$userDetails = Get-ADUser -Identity jdebla01 -Properties DisplayName, Title, Department, Company, StreetAddress, City, State, PostalCode, mail, physicalDeliveryOfficeName, Office

# Populate the $UserProperties hashtable with the required information
$UserProperties = @{
    'DisplayName'   = if ($userDetails.DisplayName) { $userDetails.DisplayName } else { "N/A" }
    'Title'         = if ($userDetails.Title) { $userDetails.Title } else { "N/A" }
    'Department'    = if ($userDetails.Department) { $userDetails.Department } else { "N/A" }
    'Company'       = if ($userDetails.Company) { $userDetails.Company } else { "N/A" }
    'Office'        = if ($userDetails.physicalDeliveryOfficeName) { $userDetails.physicalDeliveryOfficeName } else { "N/A" }
    'StreetAddress' = if ($userDetails.StreetAddress) { $userDetails.StreetAddress } else { "N/A" }
    'City'          = if ($userDetails.City) { $userDetails.City } else { "N/A" }
    'State'         = if ($userDetails.State) { $userDetails.State } else { "N/A" }
    'PostalCode'    = if ($userDetails.PostalCode) { $userDetails.PostalCode } else { "N/A" }
    'EmailAddress'  = if ($userDetails.mail) { $userDetails.mail } else { "N/A" }
}

# Specify the logo path and folder path for the signature
$logoPath = "\\dc10\c$\Windows\SYSVOL_DFSR\sysvol\tgioa.com\scripts\signatures\tgi_files\image001.jpg" # Update with the actual path to your logo
$signatureFolderPath = Join-Path -Path "C:\Users\$env:username\AppData\Roaming\Microsoft\Signatures" -ChildPath $currentUser

# Ensure the signature folder exists
if (-not (Test-Path -Path $signatureFolderPath)) {
    New-Item -ItemType Directory -Path $signatureFolderPath -Force
}

# Create the email signature
Create-EmailSignature -FolderPath $signatureFolderPath -UserProperties $UserProperties -LogoPath $logoPath

