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
11) Checks if the address is valid, if not it can create a table in the Email.

Screenshot attached below.

![Screenshot (4)](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/75174129-1ecc-430e-b990-271109959462)
![Screenshot (5)](https://github.com/MeditatingMonkey/Directory_Auditor./assets/68747956/8e2e2b1c-5a68-4076-89f2-43b409511b90)


