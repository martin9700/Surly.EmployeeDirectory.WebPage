<#
.SYNOPSIS
	Create a simple, yet good looking Employee Directory in HTML format.
.DESCRIPTION
	Run this script and create a HTML based Employee Directory that your users
	can use to locate each other's information.  Supports title, extension, cell,
	fax, description, manager, home page link and email.  Also fully utilizies Active 
    Directory's ability to store photos and you will have the choice of how you 
    want to display those pictures, either hover over or click their first name or 
    last name to see the picture (you choose).  
    
    Script includes a search box and will dynamically generate a "button bar" for
    searching for groups based on Location, Department or Manager.
	
	When the script runs, it will not pull the photo from Active Directory if
	the image file already exists in the $OutputPath\images directory.  To fully refresh
	all images (in case one is changed or removed) you just use the -Refresh parameter.
	
	Recommend the script be run hourly to keep it up to date.  Once a day I recommend
	it be run with the -Refresh parameter, which will fully refresh all images
    and delete any images from employee's no longer in Active Directory.
    
    If you need a simple application for editing this information and for uploading
    pictures to Active Directory, try my Employee Editor application:
    http://community.spiceworks.com/scripts/show/1369-employee-directory-editor
    
    ** The Office field (LDAP field name is physicalDeliveryOfficeName) is a key field for
    the script.  Put the word "Exclude" in the field and even if the object matches
    the search parameter it will not be put in the Employee Directory.  Any other
    value in here will be used as the Location display field, if you use this as your
    field to key the button bar off of, than you will get a custom button with whatever
    value you put in that field.  If the field is left blank then the script will
    assign the default location parameter into the field (only on the Employee
    Directory, Active Directory will not be updated). **
    
    Configuration is now full controlled by a separate INI file, this should make
    configuring it easier.  If ed.ini is not located in the $Path then one will be
    created, Notepad will be launched so you can edit it and the script will exit.
    
    Fields to edit in ed.ini:
    ----------------------------------------------------------------------------------------------
    HTMLTitle                   Title used for the Employee Directory
    SearchOU                    FQDN values of the OU's you want the script to search for
                                eligible users to display.  If your FQDN contains a comma
                                you must surround it with double quotes (no single quotes).
                                IE: "OU=MyCompany,DC=MyDomain,DC=local"
                                This is a search parameter, so exact matches are not necessary.
                                The same example above could be done with:
                                OU=MyCompany
    DefaultLocation             One to one match with SearchOU. Give the OU's you searched on
                                above a default name that's friendly for display.  So instead
                                of having a field on your Employee Directory that reads:
                                "OU=MyCompany,DC=MyDomain,DC=local", you can replace that with
                                "My Company, Inc.".  For every entry in SearchOU you must have
                                a matching entry in DefaultLocation.
    OutputPath                  This is the path to where you want to save the resulting HTML
                                file.  Can be either a drive letter path, or a full UNC:
                                C:\inetpub\wwwroot     or
                                \\servername\c$\inetpub\wwwroot
    UseFields                   Designate the fields you want to display on your directory. 
                                Separate the fields with comma's.  See field list below.
    SortBy                      Designate the field you want to sort by, typically LastName. Valid
                                fields to sort by are LastName, FirstName, Ext, Department, Title, 
                                Fax, Email, Location, Link, Cell and Manager
    ButtonBy                    Designate which field you want to use to create the button bar
                                across the top of the directory.  Valid choices are:  Department,
                                Manager and Location.
                                
    Fields available for Employee Directory are:
        Field                  Description
        -----                  -----------
        Picture                Small 64x64 thumbnail of the employee photo.  Hover over the
                               photo to see full sized version.
        LastName               Last Name
        FirstName              First Name
        LNLink                 Last Name as a hyperlink to user's Web Page
        LNLinkPic              Last Name as a hyperlink to user's photo
        FNLink                 First Name as a hyperlink to user's Web Page                 
        FNLinkPic              First Name as a hyperlink to user's photo
        Ext                    Active Directory field: TelephoneNumber
        Department             Department
        Title                  Title
        Fax                    Fax
        Email                  Email Address as mailto: hyperlink
        Location               Active Directory field: Office, if blank DefaultLocation will be used
        Link                   Active Directory field: Web Page
        Cell                   Active Directory field: Mobile
        Manager                First and last name of user's Manager, if set
        Description            User's description field
    
.PARAMETER Path
    Script will look for two files in this directory:  Heading.HTML and CSS.HTML.  If
    Heading.HTML is located all contents will be inserted into the heading zone of the
    Employee Directory.  Use this for custom headings, banners, images or text that you
    wish to include on your page.  Text needs to be an HTML fragment, only including
    the text you wish.  Make sure there is no HTML, HEAD, BODY or other HTML tags.
    
    If CSS.HTML is found, this will be used instead of the default CSS settings.  Text
    file should include CSS only, no HEAD or STYLE tags.
    
    If no path is set then by default the path will use the folder where the script is
    located.  Log and configuration files are also located in this path.
.PARAMETER Refresh
	If specified this parameter will clear all images from the $OutputPath\Images
	folder and freshly download everything from Active Directory.  If not specified
	it will only pull a photo from Active Directory if it does not already exist
	in the $OutputPath\Images folder.
.INPUTS
    Heading.HTML - Custom block for your HTML
    CSS.HTML - Custom CSS
    ed.ini
.OUTPUTS
	EmployeeDirectory.HTML file
    edMMDDYYhhmm.log - log file, look for errors here
.EXAMPLE
	.\Out-EmployeeDirectory.ps1
	Runs the script with all defaults
.EXAMPLE
	.\Out-EmployeeDirectory.ps1 -Path c:\EmployeeDirectory -Refresh
	Runs the script pointing all external files (input and output) to c:\EmployeeDirectory.
    Refresh has been turned on so all images will be deleted and re-downloaded 
    from Active Directory and put into c:\EmployeeDirectory\Images
.NOTES
    Author:             Martin Pugh
    Twitter:            @thesurlyadm1n
    Spiceworks:         Martin9700
    Blog:               www.thesurlyadmin.com
       
    Changelog:
        2.1             Moved all configuration settings into an INI file.  I think this will make
                        using the script a lot easier.  Also allows for changes in the code without
                        having to change parameters all the time.  Added logging.
        2.03            Found a bug in determining the HTMLPath parameter (if it's left blank it 
                        defaults to script location).
        2.02            Bug found where button wasn't filtering properly if it had a & symbol in it.
                        Also corrected a bug where an "empty" button would appear.
        2.01            Corrected bug introduced with PowerShell 3.0.  $MyInvocation.MyCommand.Path
                        no longer works in PS 3.0.  Now using Get-PSCallStack.  Tested on 3.0 and
                        2.0 and seems to work.
        2.0             Major version upgrade!  You can now choose what fields you want to display,
                        what field you want to sort on and what field you want to the button bar
                        to be based on.  There is also the capacity to select which OU's you want
                        to include in the search (Regex search).  You can also have a section for
                        custom information (imported from a file--HTML format) and you can import
                        custom CSS.  Verbose output is available during testing, if you want to
                        see it.  
        1.01            No functional change, updated comment-based help and a couple
                        of small formatting pieces here and there.  Also changed the 
                        REFRESH parameter to a switch type to make it a little 
                        friendlier to use.
        1.0             Initial Release
.LINK
	http://community.spiceworks.com/scripts/show/1652-create-employee-directory-web-page
.LINK
    http://community.spiceworks.com/scripts/show/1369-employee-directory-editor
#>
[CmdletBinding()]
Param (
	[string]$Path,
	[switch]$Refresh
)

#Functions
#region Functions
Function Set-Grid {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [object[]]$HTMLInput
    )
    Begin {
        $HTMLOutput = @()
    }
    Process {
        ForEach ($Line in $HTMLInput)
        {   
            Switch -regex ($Line)
            {   "<td>\[image\](.*?)<\/td>"
                    {   If ($Data[$Matches[1]].PicturePath)
                        {   $Line = $Line.Replace("[image]$($Matches[1])","<a onMouseOver=""ShowPicture('div$($Data[$Matches[1]].SamAccountName)')"" onMouseOut=""HidePicture('div$($Data[$Matches[1]].SamAccountName)')""><img src="".\images\$($Data[$Matches[1]].SamAccountName).jpg"" height=64 width=64></a><div class=""styleHidePicture"" id=""div$($Data[$Matches[1]].SamAccountName)""><img src="".\images\$($Data[$Matches[1]].SamAccountName).jpg""></div>")
                        }
                        Else
                        {   $Line = $Line.Replace("[image]$($Matches[1])","")
                        }
                    } #End Image
                "<td>\[email\](.*?)<\/td>"
                    {   If ($Matches[1])
                        {   $Line = $Line.Replace("[email]$($Matches[1])","<a href=""mailto:$($Matches[1])"" TITLE=""Click to E-mail Employee"">$($Matches[1])</a>")
                        }
                        Else
                        {   $Line = $Line.Replace("[email]","")
                        }
                    } #End Email
                "<td>\[fnlinkpic\](.*?)<\/td>"
                    {   If ($Data[$Matches[1]].FirstName)
                        {   $FN = $Data[$Matches[1]].FirstName
                        }
                        Else
                        {   $FN = "Link"
                        }   
                        If ($Data[$Matches[1]].PicturePath)
                        {   $Line = $Line.Replace("[fnlinkpic]$($Matches[1])","<a href=""$($Data[$Matches[1]].PicturePath)"">$FN</a>")
                        }
                        Else
                        {   $Line = $Line.Replace("[fnlinkpic]$($Matches[1])",$FN.Replace("Link",""))
                        }
                    } #End FirstName Link to Pic
                "<td>\[lnlinkpic\](.*?)<\/td>"
                    {   If ($Data[$Matches[1]].LastName)
                        {   $LN = $Data[$Matches[1]].LastName
                        }
                        Else
                        {   $LN = "Link"
                        }
                        If ($Data[$Matches[1]].PicturePath)
                        {   $Line = $Line.Replace("[lnlinkpic]$($Matches[1])","<a href=""$($Data[$Matches[1]].PicturePath)"">$LN</a>")
                        }
                        Else
                        {   $Line = $Line.Replace("[lnlinkpic]$($Matches[1])",$LN.Replace("Link",""))
                        }
                    } #End LastName Link to Pic
                "<td>\[fnlink\](.*?)<\/td>"
                    {   $DN = $Matches[1]
                        If ($Data[$DN].FirstName)
                        {   $FN = $Data[$DN].FirstName
                        }
                        Else
                        {   $FN = "Link"
                        } #End If
                        If ($Data[$DN].URL)
                        {   $Link = Set-URL $Data[$DN].URL
                            $Line = $Line.Replace("[fnlink]$DN","<a href=""$Link"" target=""_blank"">$FN</a>")
                        }
                        Else
                        {   $Line = $Line.Replace("[fnlink]$($Matches[1])",$FN.Replace("Link",""))
                        } #End If
                    } #End FirstName Link to Home Page
                "<td>\[lnlink\](.*?)<\/td>"
                    {   $DN = $Matches[1]
                        If ($Data[$DN].LastName)
                        {   $LN = $Data[$DN].LastName
                        }
                        Else
                        {   $LN = "Link"
                        } #End If
                        If ($Data[$DN].URL)
                        {   $Link = Set-URL $Data[$DN].URL
                            $Line = $Line.Replace("[lnlink]$DN","<a href=""$Link"" target=""_blank"">$LN</a>")
                        }
                        Else
                        {   $Line = $Line.Replace("[lnlink]$($Matches[1])",$LN.Replace("Link",""))
                        } #End If
                    } #End LastName Link to Home Page
                "<td>\[link\](.*?)<\/td>"
                    {   $DN = $Matches[1]
                        If ($Data[$DN].URL)
                        {   $Link = Set-URL $Data[$DN].URL
                            $Line = $Line.Replace("[link]$DN","<a href=""$Link"" TITLE=""Home Page"" target=""_blank"">Home Page</a>")
                        }
                        Else
                        {   $Line = $Line.Replace("[link]$DN","")
                        }
                    } #End Link to Home Page
            } #End Switch
            $Line = $Line.Replace("<table>","<table id='TableMain'>")
            $Line = $Line.Replace("<colgroup>","")
            $Line = $Line.Replace("<col/>","")
            $Line = $Line.Replace("</colgroup>","<CAPTION><h2>$Title</h2></CAPTION><THEAD>")
            $Line = $Line.Replace("</th></tr>","</th></TR></THEAD><TBODY>")
            $Line = $Line.Replace("</table>","")
            $Line = $Line.Replace("FirstName","First Name")
            $Line = $Line.Replace("FNLinkPic","First Name")
            $Line = $Line.Replace("FNLink","First Name")
            $Line = $Line.Replace("LastName","Last Name")
            $Line = $Line.Replace("LNLinkPic","Last Name")
            $Line = $Line.Replace("LNLink","Last Name")
            $Line = $Line.Replace("[","&#91;")
            $Line = $Line.Replace("]","&#93;")
            $HTMLOutput += $Line
        } #End ForEach
    } #End Process
    End {
        Return $HTMLOutput
    } #End End
} #End Set-Grid Function

Function Set-URL
{   Param (
        [string]$URL
    )
    If ($URL -match "https?:\/\/")
    {   $Link = $URL
    }
    Else
    {   If ($URL)
        {   $Link = "http://$URL"
        }
        Else
        {   $Link = $null
        }
    } #End If
    Return $Link
} #End Set-URL Function

Function Set-EscapeCharacters
{   Param (
        [string]$Field
    )
    $Field = $Field.Replace("[","&#91;")
} #End Set-EscapeCharacters Function

Function Convert-CSVData
{   Param (
        [string]$CSV
    )
    
    $Data = @()
    Do {
        If ($CSV.Substring(0,1) -eq ",")
        {   $CSV = $CSV.Substring(1)
        }
        $Value = ($CSV | ConvertFrom-Csv -Header Field).Field
        If ($CSV -ne "")
        {   If ($CSV.Substring(0,1) -eq """")
            {   $CSV = $CSV.Substring(($Value.Length + 2))
            }
            Else
            {   $CSV = $CSV.Substring($Value.Length)
            }
        }
        $Data += $Value
    } Until ($CSV -eq "")
    Return $Data
} #End Convert-CSVData Function

#endregion

#Check if paths exist
If ($Path)
{   If (-not (Test-Path $Path -PathType Container))
    {   Write-Verbose "$(Get-Date): $Path doesn't exist, defaulting to script location"
        $HTMLPath = Split-Path $MyInvocation.MyCommand.Path
    }
    Else
    {   $HTMLPath = $Path
    }
}
Else
{   $HTMLPath = Split-Path $MyInvocation.MyCommand.Path
}
$Today = Get-Date -Format MMddyyyyhhmm

Start-Transcript "$HTMLPath\ed$Today.log"

#Setup
Write-Verbose "$(Get-Date): Script begins"
Write-Verbose "$(Get-Date): Parsing INI file"
If (-not (Test-Path $HTMLPath\ed.ini -Pathtype Leaf))
{   Write-Host "Configuration file not present, creating.  Please edit and rerun the script." -ForegroundColor Red
    $edini = @"
[SETUP]
HTMLTitle=Spiceworks Employee Directory

[SEARCH]
SearchOU="OU=Boston","OU=NewYork","OU=Gaithersburg","OU=Faxes"
DefaultLocation=Boston,New York,Gaithersburg,Boston

[PATHS]
OutputPath=\\server\wwwshare

[FIELDS]
UseFields=Picture,LastName,FirstName,Title,Manager,Location,Ext,Cell,Fax,Email
SortBy=LastName
ButtonBy=Location
"@
    $edini | Out-File $HTMLPath\ed.ini
    Notepad.exe $HTMLPath\ed.ini
    Exit
}

$edini = Get-Content $HTMLPath\ed.ini
$Title = ($edini | Select-String "HTMLTitle").Line.SubString(10)
$SearchOU = ($edini | Select-String "SearchOU").Line.SubString(9)
$LocationDefault = Convert-CSVData -CSV (($edini | Select-String "DefaultLocation").Line.SubString(16))
$OutputPath = ($edini | Select-String "OutputPath").Line.SubString(11)
$Fields = Convert-CSVData -CSV (($edini | Select-String "UseFields").Line.SubString(10))
$SortBy = ($edini | Select-String "SortBy").Line.SubString(7)
$ButtonBy = ($edini | Select-String "ButtonBy").Line.SubString(9)
$ImagesPath = "$OutputPath\images"

#Process SearchOU
$Domains = @{}
ForEach ($Line in $SearchOU)
{
    If ($Line -like "*:*")
    {
        $Info = $Line.Split(":")
        $Domains.Add($Info[0],(Convert-CSVData -CSV $Info[1]))
    }
    Else
    {
        $Domains.Add("",(Convert-CSVData -CSV $Line))
    }
}


If (-not (Test-Path $OutputPath -PathType Container))
{	Write-Error "`nPath $OutputPath does not exist, script cannot continue."
	Exit
}
Else
{	If (-not (Test-Path $ImagesPath))
	{	Write-Verbose "$(Get-Date): $ImagesPath not detected, attempting to create"
        Try {
            New-Item -Path $ImagesPath -ItemType Directory -ErrorAction Stop | Out-Null
        }
        Catch {
            Write-Error $Error[0]
            Exit
        }
	}
}

#Verify all field entries are good
$Regex = "Picture|LastName|FirstName|LNLink|LNLinkPic|FNLink|FNLinkPic|Ext|Department|Title|Fax|Email|Location|Link|Cell|Manager|Description"
ForEach ($Field in $Fields.Split(","))
{   If ($Field -notmatch $Regex)
    {   Write-Error "Field: $Field, is not a valid field"
        Exit
    }
}

#Determine Filter Column
If (-not ($ButtonBy -match "Department|Location|Manager"))
{   Write-Error "ButtonBy must be either Department, Location or Manager.  Set to: $ButtonBy"
    Exit
}
Else
{   $FilterColumn = 0
    ForEach ($Column in $Fields.Split(","))
    {   If ($Column -eq $ButtonBy)
        {   $Found = $true
            Break
        }
        $FilterColumn ++
    }
    If (-not $Found)
    {   Write-Error "ButtonBy parameter does not match a specified field"
        Exit
    }
}

#Verify Good Sort Field
$Regex = $Fields.Replace(",","|") + "|LastName|FirstName"
If ($SortBy -notmatch $Regex)
{   Write-Error "SortBy parameter does not match a specified field in `$Fields"
    Exit
}

#Change the Search parameter to Regex
[string[]]$Temp = $Domains.Keys
$SearchRegexDomain = $Temp -join "|"
$SearchRegexDomain = $SearchRegexDomain.Replace("/","\/")  #Escape some likely characters
$SearchRegexDomain = $SearchRegexDomain.Replace(".","\.")

Write-Verbose "$(Get-Date):      Search: $SearchRegexOU"
Write-Verbose "$(Get-Date): Output Path: $OutputPath"
Write-Verbose "$(Get-Date): Images Path: $ImagesPath"

#Refresh images?
If ($Refresh)
{   Write-Verbose "$(Get-Date): Full refresh requested, deleting old images..."
    Try {
        Remove-Item $ImagesPath\*.jpg -Force -ErrorAction Stop
    }
    Catch {
        Write-Error $Error[0]
        Exit
    }
}

#Custom heading?
If (Test-Path $HTMLPath\Heading.HTML)
{   Write-Verbose "$(Get-Date): Custom heading detected $HTMLPath\Heading.HTML adding to Employee Directory"
    $HeadingHTML = Get-Content $HTMLPath\Heading.HTML
}
Else
{   Write-Verbose "$(Get-Date): Custom $HeadingPath\Heading.HTML not found, will not be included in Employee Directory"
}

#Custom CSS?
If (Test-Path $HTMLPath\CSS.HTML)
{   Write-Verbose "$(Get-Date): Custom CSS detected $HTMLPath\CSS.HTML overriding default CSS"
    $CSSHTML = Get-Content $HTMLPath\CSS.HTML
}
Else
{   Write-Verbose "$(Get-Date): No Custom CSS detected, using default"
    #Define Default CSS
    $CSSHTML = @"

form { 
  margin: 0; 
} 
table {
 background:#D3E4E5; 
 border:1px solid gray; 
 border-collapse:collapse; 
 color:#fff; 
 font:normal 12px verdana, arial, helvetica, sans-serif; 
 width:95%;
} 
caption { border:1px solid #5C443A; 
 color:#5C443A; 
 font-weight:bold; 
 padding:6px 4px 8px 0px; 
 text-align:center; 
} 
td, th { color:#363636; 
 padding:.4em; 
} 
tr { border:1px dotted gray; 
} 
thead th, tfoot th { background:#5C443A; 
 color:#FFFFFF; 
 padding:3px 10px 3px 10px; 
 text-align:left; 
 text-transform:uppercase; 
} 
tbody th, tbody td { text-align:left; 
 vertical-align:top; 
} 
tbody tr:hover { background:#99BCBF; 
 border:1px solid #03476F; 
 color:#000000; 
} 
"@ #End Default CSS
} #End Custom or Default CSS

#Additional CSS
$CSSHTML += @"

.styleHidePicture {
position:absolute;
visibility:hidden;
}
.styleShowPicture {
position:absolute;
visibility:visible;
border:solid 7px Black;
padding:1px;
}

"@

$JSHTML = @"
<BODY>
<script type='text/javascript'>
function filter (phrase, _id){
   var words = phrase.value.toLowerCase().split(" ");
   var table = document.getElementById(_id);
   var ele;
   for (var r = 1; r < table.rows.length; r++){
         ele = table.rows[r].innerHTML.replace(/<[^>]+>/g,"");
           var displayStyle = 'none';
           for (var i = 0; i < words.length; i++) {
             if (ele.toLowerCase().indexOf(words[i])>=0)
               displayStyle = '';
             else {
               displayStyle = 'none';
               break;
             }
           }
         table.rows[r].style.display = displayStyle;
   }
}
function filtbutton(phrase){
    var tableMain = document.getElementById('TableMain');
	   for(i=1;i<tableMain.rows.length;i++){
        ele = tableMain.rows[i].innerHTML.replace(/<[^>]+>/g,"");
        tableMain.rows[i].style.display = '';}
    for(i=1;i<tableMain.rows.length;i++){
	       ele = tableMain.rows[i].cells[$FilterColumn].innerHTML.replace(/<[^>]+>/g,"");
           ele = ele.replace("&amp;","\&")
        if (ele != phrase && phrase != '') {
            tableMain.rows[i].style.display = 'none';
        } else {
        }
    }
}
<!--
function ShowPicture(id) {
	var currentDiv = document.getElementById(id);
	currentDiv.className='styleShowPicture'
 }
 function HidePicture(id) {
	var currentDiv = document.getElementById(id);
	currentDiv.className='styleHidePicture'
 }
 //-->
</script>

"@ #End JSHTML

$HeaderHTML = @"
<HTML>
<HEAD>
  <TITLE>$Title</TITLE>
<style type='text/css'>

"@  #End HeaderHTML

$EndHeaderHTML = @"
</style>
</HEAD>

"@ #End EndHeaderHTML

$SearchHTML = @"
</tr></table><table id='Filterline'><TD><div align='right'><FORM><FONT face='verdana,arial,helvetica,sans-serif' size=2><b>Search: </b></FONT><input name='filt' onkeyup="filter(this, 'TableMain', '1')" type='text'></FORM></div><br></TD>
</TABLE>

"@ #End SearchHTML

$FooterHTML = @"
</TBODY>
</TABLE>
</BODY>
</HTML>
"@ #End FooterHTML

#Get User Information load it into an object for later
Write-Verbose "$(Get-Date): Gathering data from Active Directory"
$Users = ForEach ($Key in $Domains.Keys)
{
    If ($Key)
    {
        $GC = "GC://$Key"
    }
    Else
    {
        $GC = ""
    }
    $ADSearch = New-Object System.DirectoryServices.DirectorySearcher
    $ADSearch.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry($GC)
    $ADSearch.SearchScope = "Subtree"
    $ADSearch.Filter = "(objectCategory=User)"
    $PropertiesToLoad = "SamAccountName,useraccountcontrol,distinguishedname,GivenName,sn,Title,description,department,physicaldeliveryofficename,manager,TelephoneNumber,Mobile,facsimiletelephonenumber,mail,thumbnailphoto,wwwhomepage,description"
    ForEach ($Property in $($PropertiesToLoad.Split(",")))
    {	$ADSearch.PropertiesToLoad.Add($Property) | Out-Null
    }
    $ADSearch.FindAll()
}

$Data = @{}
ForEach ($User in $Users)
{	#Filter out users who don't fall within the search parameters
    $Found = $false
    ForEach ($Key in $Domains.Keys)
    {
        $SearchRegexOU = $Domains[$Key] -join "|"
        $SearchRegexOU = $SearchRegexOU.Replace("/","\/")  #Escape some likely characters
        $SearchRegexOU = $SearchRegexOU.Replace(".","\.")

        If ($($User.Properties.distinguishedname) -match $SearchRegexDomain -and $($User.Properties.distinguishedname) -match $SearchRegexOU)
        {   
            $Found = $true
            Break
        }
    }

    If (-not $Found)
    {
        Continue
    }

    #Filter out users with the word 'Exclude' in the department field
	If ($($User.Properties.physicaldeliveryofficename))
	{	If ($($User.Properties.physicaldeliveryofficename).ToUpper() -eq "EXCLUDE")
		{	Continue
		}
	}
    
	#Filter out any disabled users
	If ($($User.Properties.useraccountcontrol) -band 0x2)
	{	Continue
	}
    
    #Retrieve the Picture
	$File = "$ImagesPath\$($User.Properties.samaccountname).jpg"
	If (-not (Test-Path $File))
	{	If (($User.Properties.thumbnailphoto).Count)
		{	Try {
                $User.Properties.thumbnailphoto | Set-Content -Path $File  -Encoding Byte -ErrorAction Stop
            }
            Catch {
                Write-Error "Unable to save thumbnail to $ImagesPath, see above error.  Aborting script."
                Break
            }
		}
        Else
        {   $File = $null
        }
	} 
    
    #Now load the data
	$Object = New-Object PSObject -Property @{
		LastName = $($User.Properties.sn)
		FirstName = $($User.Properties.givenname)
		Title = $($User.Properties.title)
		Department = $($User.Properties.department)
        Location = $($User.Properties.physicaldeliveryofficename)
        Manager = $($User.Properties.manager)
		Ext = $($User.Properties.telephonenumber)
		Cell = $($User.Properties.mobile)
		Fax = $($User.Properties.facsimiletelephonenumber)
		Email = "[email]$($User.Properties.mail)"
		SamAccountName = $($User.Properties.samaccountname)
        Picture = "[image]$($User.Properties.distinguishedname)"
        PicturePath = $File
        URL = $($User.Properties.wwwhomepage)
        Link = "[link]$($User.Properties.distinguishedname)"
        FNLinkPic = "[fnlinkpic]$($User.Properties.distinguishedname)"
        LNLinkPic = "[lnlinkpic]$($User.Properties.distinguishedname)"
        FNLink = "[fnlink]$($User.Properties.distinguishedname)"
        LNLink = "[lnlink]$($User.Properties.distinguishedname)"
        Description = $($User.Properties.description)
	}
	$Data.Add($($User.Properties.distinguishedname),$Object)
}

#Populate Manager and Default Location
ForEach ($User in $Data.Values)
{   If ($User.Manager)
    {   $User.Manager = "$($Data[$User.Manager].FirstName) $($Data[$User.Manager].LastName)"
    }
    If ($User.Location)
    {   $User.Location = $User.Location.Trim()
    }
    If ([string]::IsNullOrEmpty($User.Location))
    {   For ($i = 0;$i -le ($SearchOU.Count - 1);$i++)
        {   If ($User.Link -like "*$($SearchOU[$i])*")
            {   $User.Location = $LocationDefault[$i]
                Break
            }
        }
    }
} #End ForEach

#Build the Button row
Write-Verbose "$(Get-Date): Building the HTML"
$Buttons = $Data.Values | Where { $_ } | Select $ButtonBy -Unique | Sort $ButtonBy
$ButtonHTML = @"
<TABLE id='Filterbutton'>
<TD><input id="All" value="All" onClick="filtbutton('')" type="button"/></TD>

"@
$MaxButtons = 2
ForEach ($Button in $Buttons)
{	If ($Button.$ButtonBy)
	{	$MaxButtons ++
        $ButtonHTML += "<TD><input id=""$($Button.$ButtonBy)"" value=""$($Button.$ButtonBy)"" onClick=""filtbutton('$($Button.$ButtonBy)')"" type=""button""/></TD>`n"
        If ($MaxButtons -eq 9)
        {   $MaxButtons = 1
            $ButtonHTML += "</tr><tr>`n"
        }
	}
}

#Build the detail HTML
$GridHTML = $Data.Values | Sort $SortBy | Select $($Fields.Split(",")) | ConvertTo-Html -Fragment | Set-Grid

#Put it together and save
$HeadHTML = $HeaderHTML + $CSSHTML + $EndHeaderHTML + $JSHTML
$FullHTML = $HeadHTML + $HeadingHTML + $ButtonHTML  + $SearchHTML + $GridHTML + $FooterHTML
Write-Verbose "$(Get-Date): Saving HTML: $OutputPath\EmployeeDirectory.html"
$FullHTML | Out-File $OutputPath\EmployeeDirectory.html
#& $OutputPath\EmployeeDirectory.html                             #Un-remark if you wish to have the page displayed automatically in your browser
Write-Verbose "$(Get-Date): Script completed"

Stop-Transcript

#Clean up old log files
Get-ChildItem "$HTMLPath\ed??????????.log" | Where { $_.LastWriteTime -le (Get-Date).AddDays(-30).Date } | Remove-Item