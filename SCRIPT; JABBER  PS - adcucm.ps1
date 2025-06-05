# Specify the path to your CSV file containing the phone numbers and user information
$csvPath = "C:\Users\XXXXX\Desktop\CUCM_Automation_Files\Jabber Automation\user_data.csv"

# Specify the column names in your CSV file
$csvColumns = "Username", "PhoneNumber"

# Function to update the telephoneNumber attribute for a user
Function Update-UserPhoneNumber {
    param (
        [string]$username,
        [string]$phoneNumber
    )
    
    # Get the user by username
    $user = Get-ADUser -Filter {SamAccountName -eq $username}
    
    if ($user) {
        # Update the telephoneNumber attribute
        Set-ADUser -Identity $user -Replace @{telephoneNumber=$phoneNumber}
        Write-Host "Updated phone number for $username to $phoneNumber"
    }
    else {
        Write-Host "User $username not found in Active Directory."
    }
}

# Read the CSV file and update user phone numbers
Import-Csv -Path $csvPath -Header $csvColumns | ForEach-Object {
    Update-UserPhoneNumber -username $_.Username -phoneNumber $_.PhoneNumber
}