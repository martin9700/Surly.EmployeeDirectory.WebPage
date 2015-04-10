For more detailed instructions, see here: 
http://thesurlyadmin.com/2013/05/16/new-version-employee-directory/

Run this script and create a HTML based Employee Directory that your users can use to locate each other's information. Supports title, extension, cell, fax, description, manager, home page link and email. Also fully utilizies Active Directory's ability to store photos and you will have the choice of how you want to display those pictures, either hover over or click their first name or last name to see the picture (you choose). 

Script includes a search box and will dynamically generate a "button bar" for searching for groups based on Location, Department or Manager. 

When the script runs, it will not pull the photo from Active Directory if the image file already exists in the $OutputPath\images directory. To fully refresh all images (in case one is changed or removed) you just use the -Refresh parameter. 

Recommend the script be run hourly to keep it up to date. Once a day I recommend it be run with the -Refresh parameter, which will fully refresh all images and delete any images from employee's no longer in Active Directory.

1. Download the script and save it as Out-EmployeeDirectory.ps1 
2. Edit parameters to match your needs 
3. Setup a scheduled task to run on an hourly basis. 
3a. Run Powershell.exe 
3b. As arguments use: -ExecutionPolicy Bypass pathtoscript\Out-EmployeeDirectory.ps1 
4. Setup a second scheduled task to run once a day, use this as the arguments: 
4a. -ExecutionPolicy Bypass pathtoscript\Out-EmployeeDirectory.ps1 -Refresh

For more information on setting up Powershell in a scheduled task: http://community.spiceworks.com/how_to/show/17736-run-powershell-scripts-from-task-scheduler

For additional help on the script, open a Powershell prompt at type:

Get-Help pathtoscript\Out-EmployeeDirectory.ps1 -Full

Looking for an easy way to edit this data? Maybe even delegate it to your HR department? 
http://community.spiceworks.com/scripts/show/1369-employee-directory-editor

Changelog: 
2.02 Bug found where button wasn't filtering properly if it had a & symbol in it. Also corrected 
a bug where an "empty" button would appear. 
2.0 Major version upgrade! You can now choose what fields you want to display, 
what field you want to sort on and what field you want to the button bar 
to be based on. There is also the capacity to select which OU's you want 
to include in the search (Regex search). You can also have a section for 
custom information (imported from a file--HTML format) and you can import 
custom CSS. Verbose output is available during testing, if you want to 
see it. 
1.01 No functional change, updated comment-based help and a couple 
of small formatting pieces here and there. Also changed the 
REFRESH parameter to a switch type to make it a little 
friendlier to use. 
1.0 Initial Release
