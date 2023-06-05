# Directory_Auditor.
Description
This PowerShell script provides an efficient way to enumerate all the users from various groups who have access to a particular directory. This script is incredibly useful for system administrators, security analysts, and those concerned with access management, especially in complex environments where numerous groups and nested groups can have varying levels of access. This script can take multiple directory addresses at once. For each address it creates a new sheet in the excel file. 

Features:
1) Retrieves all users who have access to a specific directory through group memberships
2) Handles nested groups, so you won't miss a single user
3) Lists users from both local and domain groups
4) Can be customized to target different directories or domains
5) Outputs a convenient and easy-to-read list of users
6) Helps to streamline audits of folder permissions
7) Provides a visual report of who exactly has what access
8) Shows the permission (Read/Write , Read Only) of a user.
9) Sends the Excel file to the email directly.
10) We can add cc recipient to the script.
