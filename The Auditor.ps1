##################################################
##
##The Auditor
##
##Description:
##This Script is pulling all the users from the Groups having an access to the Directory
##
##Make Sure to Change the Directory Path for the results.
##
##By : Tusshar Singh
##
##Date: 16 May 2023
##
##################################################
# Import the Active Directory module and ImportExcel module
Import-Module ActiveDirectory
Import-Module ImportExcel

# Define the directory paths
$directoryPaths = @("G:\Generation")

# Define the OU to exclude
$excludeOU = "OU=Service,OU=Electric,DC=corp,DC=local"

# Define the output file
$outputFile = "C:\Users\$Env:UserName\Desktop\Group Members.xlsx"

#Email Server and address parameters
$EmailFrom = "fbctechnicalsupport@fortisbc.com"
$EmailTo = "abc@xyz.com"
$cc = "klm@xyz.com"
$subject = "Monthly Auditing of Group Members"
$body = "Hello, <br><br>Kindly check the Excel File to view the members of the groups that have access to the directories."

$notfoundDirectories = @()


$SMTPServer = "smtp.corp.local"
$SMTPPort = 25

# Define a function to create the output object
function CreateOutputObject($Identity, $Rights, $DisplayName) {
    # Split FileSystemRights at comma and take the first part
    $fileSystemRights = $Rights.ToString().Split(',')[0]

    if ($fileSystemRights -eq "FullControl") {
        $fileSystemRights = "Read\Write"
    }
    elseif ($fileSystemRights -eq "Modify") {
        $fileSystemRights = "Read\Write"
    }
    elseif ($fileSystemRights -eq "ReadAndExecute") {
        $fileSystemRights = "Read Only"
    }
    else {
        return
    }

    # Create custom object to store output in a single row
    $output = New-Object -TypeName PSCustomObject
    $output | Add-Member -MemberType NoteProperty -Name 'Groups/Users' -Value $Identity
    $output | Add-Member -MemberType NoteProperty -Name 'Permission' -Value $fileSystemRights
    $output | Add-Member -MemberType NoteProperty -Name 'Users' -Value $DisplayName

    return $output
}

# Iterate over each directory
foreach($directoryPath in $directoryPaths) {
    # Create an array to store the output objects
    $outputArray = @()

    # Check if the directory exists
    if(Test-Path -Path $directoryPath) {
        # Get the Access Control List (ACL) for the directory
        $acl = Get-Acl -Path $directoryPath

        # Create a hashtable to store the ACL entries
        $aclTable = @{}

        # Iterate over each Access rule in the ACL
        foreach ($accessRule in $acl.Access) {
            # Create a key for the hashtable
            $key = "$($accessRule.IdentityReference),$($accessRule.FileSystemRights),$($accessRule.AccessControlType)"

            # Check if the key exists in the hashtable
            if ($aclTable.ContainsKey($key)) {
                # If the key exists, it means this is a duplicate ACL entry
                # Remove the duplicate ACL entry
                $acl.RemoveAccessRule($accessRule)
            } else {
                # If the key doesn't exist, add it to the hashtable
                $aclTable.Add($key, $accessRule)

                if ($accessRule.IdentityReference.ToString().Contains('\')) {
                    $identityName = $accessRule.IdentityReference.ToString().Split('\')[1]
                    $domainName = $accessRule.IdentityReference.ToString().Split('\')[0]
                } else {
                    $identityName = $accessRule.IdentityReference.ToString()
                    $domainName = $null
                }
            
                    # Check if the Identity is a group
                    $group = Get-ADGroup -Filter { Name -eq $identityName } -ErrorAction SilentlyContinue


                    if ($group) {
                        try {
                            $groupMembers = Get-ADGroupMember -Identity $group -Recursive -ErrorAction Stop
                        }
                        catch {
                            $outputArray += New-Object -TypeName PSCustomObject -Property @{
                                'Groups/Users' = "`"$($group.SamAccountName)`""
                                'Permission' = 'N/A'
                                'Users' = 'N/A'
                            }
                            continue
                        }

                        # Iterate over each member of the group
                        foreach ($member in $groupMembers) {
                            # Skip users whose name start with RL_ or who are in the Service OU
                            if ($member.SamAccountName -ne "Authenticated Users" -and $member.SamAccountName -ne "INTERACTIVE" -and $member.SamAccountName -notlike "RL_*" -and $member.DistinguishedName -notlike "*$excludeOU*") {
                               try {
                                    $user = Get-ADUser -Identity $member.SamAccountName -Properties DisplayName -ErrorAction Stop
                                }
                                catch {
                                    $outputArray += New-Object -TypeName PSCustomObject -Property @{
                                        'Groups/Users' = "`"$($member.SamAccountName)`""
                                        'Permission' = 'N/A'
                                        'Users' = 'N/A'
                                    }
                                    continue
                                }

                                # Add the user to the output
                                $outputArray += CreateOutputObject -Identity $accessRule.IdentityReference -Rights $accessRule.FileSystemRights -DisplayName $user.DisplayName
                            }
                        }
                    }
                    else {
                        # If it's not a group, it must be a user
                        # Skip users whose name start with RL_ or who are in the Service OU
                        if ($member.SamAccountName -notlike "RL_*" -and $member.DistinguishedName -notlike "*$excludeOU*") {
                    
                           if ($domainName -eq $null) {
                                $user = Get-LocalUser -Name $identityName -ErrorAction SilentlyContinue
                            } else {
                                try { 
                                    # Get the DisplayName of the user
                                    $user = Get-ADUser -Filter {SamAccountName -eq $identityName} -Properties DisplayName -ErrorAction Stop
                                } catch {
                                    $outputArray += New-Object -TypeName PSCustomObject -Property @{ 
                                    'Groups/Users' = "`"$($identityName)`""
                                    'Permission' = 'N/A'
                                    'Users' = 'N/A'
                                    }
                                    continue 
                                }

                            }
                    
                            # Add the user to the output
                            if ($user) {
                                $outputArray += CreateOutputObject -Identity $accessRule.IdentityReference -Rights $accessRule.FileSystemRights -DisplayName $user.DisplayName
                            }
                        }
                    }
                }           
            }

            # Export the array to an Excel file
            $worksheetName = Split-Path $directoryPath -Leaf
            $outputArray | Export-Excel -Path $outputFile -WorkSheetname $worksheetName -TableName $worksheetName -TableStyle Medium14 -AutoSize -Title $directoryPath -TitleSize 18
                
            Write-Output ("The worksheet '{0}' in Excel file '{1}' has been created." -f $worksheetName, $outputFile)

        }
        else {
            Write-Output ("The directory '{0}' does not exist." -f $directoryPath)

            $notfoundDirectories += $directoryPath
        }
}

if($notfoundDirectories.Length -gt 0) {
    $bodynotfound = "Hello, <br><br>Kindly check the Excel File to view the members of the groups that have access to the directories, unfortunately couldn't locate the directories mentioned below: <br><br>"
    $bodynotfound += "<table style='border: 1px solid black; border-collapse: collapse;'><tr><th bgcolor='#7E7E7E' style='border: 1px solid black; padding: 5px;'>Directories Not Found</th></tr>"
    
    foreach($dir in $notfoundDirectories) {
        $bodynotfound += "<tr><td style = 'border: 1px solid black; padding: 5px;'>$dir</td></tr>"
    }
    
    $body +="</table>"
    Send-MailMessage -SmtpServer $SMTPServer -Port $SMTPPort -From $EmailFrom -To $EmailTo -Cc $cc -Subject $subject -Body $bodynotfound -BodyAsHtml -Attachments $outputFile
} else {
    Send-MailMessage -SmtpServer $SMTPServer -Port $SMTPPort -From $EmailFrom -To $EmailTo -Cc $cc -Subject $subject -Body $body -BodyAsHtml -Attachments $outputFile
}