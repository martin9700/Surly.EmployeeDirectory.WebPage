<#
.SYNOPSIS
	Script to create an employee directory in Excel format and email it
.DESCRIPTION
	This script will create an employee directory in Excel format and email that
	file.  
	
	I would recommend setting this script to run once a month and have it email
	a distribution list of people who are interested in receiving it.
	
	Edit all the parameter default settings to match your environment.

    ** PRE-REQUISITE **
    You must have Microsoft Excel installed on the computer running the script.

.PARAMETER ExportPath
	The path where you would like the spreadsheet saved.  The spreadsheet will 
	be called empdir.xlsx.

    If no path is given the script will default to the path where the script is located.

.PARAMETER Title
	The title to be used in the spreadsheet.

.PARAMETER SearchBase
	Maybe the most important parameter for this script.  This sets the base level
	of where the script will search.  This will allow you to filter out any user
	objects that you don't want included in the script (administrator, etc).
	
	Use the FQDN of the OU where you want the search to begin--the script will 
	automatically search all OU's under the base one.

.OUTPUT
	Excel spreadsheet named empdir.xlsx in the $ExportPath folder.
.EXAMPLE
	.\Export-EmployeeDirectory.ps1 -ExportPath \\server\share\path -Title "MyCompany Employee Directory" -SearchBase "OU=Users,OU=mydomain,DC=local"
	
	Custom settings for all parameters.

.EXAMPLE
	.\Export-EmployeeDirectory.ps1
	
	Accept all defaults.  Will pull all user objects in Active Directory, including service accounts!  

.NOTES
    Author:             Martin Pugh
    Twitter:            @thesurlyadm1n
    Spiceworks:         Martin9700
    Blog:               www.thesurlyadmin.com
      
    Changelog:
        1.3             Removed SMTP functionality (not everyone will want that). Added some better path error catching, 
                        updated comments and made script flow a little better.
        1.2             Added sort on last name.  Set script to erase the file if it
                        already exists.
        1.1             Completely revamped to use ADSI instead of RSAT.  Now supports
                        user objects and contact objects.
        1.0             Initial Release

.LINK
    http://community.spiceworks.com/scripts/show/1630-export-employee-directory-to-excel-and-email
.LINK
	http://community.spiceworks.com/scripts/show/1002-employee-directory-with-photo-s
#>
Param (
	[string]$ExportPath,
	[string]$Title = "Surly Admin Employee Directory",
    [string]$SearchBase
)
	
#Set Path
If (-not $ExportPath)
{
    $Path = Split-Path $MyInvocation.MyCommand.Path
}
ElseIf (-not (Test-Path $ExportPath))
{
    Throw "Unable to locate the Export Path: $ExportPath"
}

#Get User Information load it into an object for later
$Domain = New-Object System.DirectoryServices.DirectoryEntry("")
$ADSearch = New-Object System.DirectoryServices.DirectorySearcher
$ADSearch.SearchRoot = $Domain
$ADSearch.SearchScope = "Subtree"
$ADSearch.Filter = "(objectCategory=User)"
$PropertiesToLoad = "distinguishedname,useraccountcontrol,GivenName,sn,title,department,TelephoneNumber,Mobile,facsimiletelephonenumber,mail"
ForEach ($Property in $($PropertiesToLoad.Split(",")))
{	$ADSearch.PropertiesToLoad.Add($Property) | Out-Null
}
$Users = $ADSearch.FindAll()

$Data = ForEach ($User in $Users)
{	If ($User.Properties.distinguishedname -like "*$SearchBase*")
    {   If (-not ($($User.Properties.useraccountcontrol) -band 0x2))  #Exclude disabled users
    	{	New-Object PSObject -Property @{
                sn = $($User.Properties.sn)
                Givenname = $($User.Properties.givenname)
        		Title = $($User.Properties.title)
        		Department = $($User.Properties.department)
        		Telephonenumber = $($User.Properties.telephonenumber)
        		Mobile = $($User.Properties.mobile)
        		Fax = $($User.Properties.facsimiletelephonenumber)
        		Mail = $($User.Properties.mail)
            }
    	}
    }
}


#If there's data, make the spreadsheet
If ($Data)
{
    #Excel Constants
    $xlHAlignCenterAcrossSelection = 7

    #Setup the spreadsheet
    $Excel = New-Object -ComObject Excel.Application
    #$Excel.Visible = $true    #can unremark for testing
    $Workbooks = $Excel.Workbooks.Add()
    $Worksheets = $Workbooks.Worksheets
    $Worksheet = $Worksheets.Item(1)
    $Worksheet.Name = $Title

    #Create the Title line
    $Worksheet.Cells.Item(1,1) = $Title
    $Worksheet.Cells.Item(1,1).Font.Bold = $true
    $Worksheet.Cells.Item(1,1).Font.Size = 18
    $Range = $Worksheet.Range("A1:H1")
    $Range.HorizontalAlignment = $xlHAlignCenterAcrossSelection

    #Create the Header row
    $Worksheet.Cells.Item(2,1) = "Last Name"
    $Worksheet.Cells.Item(2,2) = "First Name"
    $Worksheet.Cells.Item(2,3) = "Title"
    $Worksheet.Cells.Item(2,4) = "Department"
    $Worksheet.Cells.Item(2,5) = "Ext"
    $Worksheet.Cells.Item(2,6) = "Cell"
    $Worksheet.Cells.Item(2,7) = "Fax"
    $Worksheet.Cells.Item(2,8) = "Email"

    #Pretty up the header row a little
    1..8 | ForEach {
	    $Worksheet.Cells.Item(2,$_).Font.Bold = $true
	    $Worksheet.Cells.Item(2,$_).Font.Underline = $true
    }

    #Get users from AD and populate the spreadsheet
    $Cell = 3


    ForEach ($User in ($Data | Sort sn))
    {   $Worksheet.Cells.Item($Cell,1) = $User.sn
	    $Worksheet.Cells.Item($Cell,2) = $User.Givenname
	    $Worksheet.Cells.Item($Cell,3) = $User.Title
	    $Worksheet.Cells.Item($Cell,4) = $User.Department
	    $Worksheet.Cells.Item($Cell,5) = $User.Telephonenumber
	    $Worksheet.Cells.Item($Cell,6) = $User.Mobile
	    $Worksheet.Cells.Item($Cell,7) = $User.Fax
	    $Worksheet.Cells.Item($Cell,8) = $User.Mail
	    $Cell ++
    }

    $Cell ++
    $Worksheet.Cells.Item($Cell,1) = "Created on $(Get-Date -Format 'MMMM dd, yyyy')"
    $Worksheet.Cells.Item($Cell,1).Font.Italic = $true

    #Fix the formatting a little
    $Range = $Worksheet.UsedRange
    [Void]$Range.EntireColumn.AutoFit()

    #Save the file
    $File = Join-Path -Path $ExportPath -ChildPath "EmployeeDirectory.xlsx"
    If (Test-Path $File)
    {   Remove-Item $File -Force
    }
    $Workbooks.SaveAs($File)
    $Excel.Quit()
}
Else
{
    Write-Warning "No users in $SearchBase were located"
}
