#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ScriptDirectory
{
<#
	.SYNOPSIS
		Get-ScriptDirectory returns the proper location of the script.

	.OUTPUTS
		System.String
	
	.NOTES
		Returns the correct path within a packaged executable.
#>
	[OutputType([string])]
	param ()
	if ($null -ne $hostinvocation)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

function Connect-eduSTARMC
{
    <#
        .SYNOPSIS
            Connects to the eduSTAR Management Console

        .PARAMETER Credential
            PSCredential object containing the username and password of an account with permission to login to the eduSTAR Management Console

        .OUTPUTS
            Returns status code and description

        .EXAMPLE
            Connect-eduSTARMC

            Connect-eduSTARMC -Credential $MyCredential
    #>
	
	param (
		[pscredential]$Credential = (Get-Credential)
	)
	
	$WebRequestBody = @{
		curl	 = "Z2Fedustarmc"
		username = $Credential.UserName
		password = $Credential.GetNetworkCredential().Password
		SubmitCreds = "Log+in"
	}
	
	$Request = Invoke-WebRequest -Uri https://apps.edustar.vic.edu.au/CookieAuth.dll?Logon -Body $WebRequestBody -Method Post -SessionVariable session
	
	# This session will be reused by other functions after connecting
	$script:eduSTARMCSession = $session
	
	# If the connection is open, get information about the currently logged in user
	if ($Request.Headers.Connection -eq 'Keep-Alive')
	{
		$GetUser = Invoke-RestMethod -Uri https://apps.edustar.vic.edu.au/edustarmc/api/MC/GetUser -Method Get -WebSession $eduSTARMCSession -ContentType "application/xml"
		
		$obj = New-Object psobject -Property ([ordered]@{
				Connected  = $true
				LoggedInAs = $GetUser.User._displayName
				UserDetails = $GetUser.User
				Schools    = $GetUser.User._schools.ChildNodes.Count
			})
		
		$script:eduSTARMCConnection = $obj
		
		Write-Output $obj
		
	}
	else
	{
		throw 'Unable to connect to the eduSTAR Management Console'
	}
}

#region Schools

function Get-eduSTARMCSchool
{
    <#
        .SYNOPSIS
            Returns all schools assigned to the user currently authenticated to the eduSTAR Management Console

        .OUTPUTS
            Returns SchoolId, SchoolName and Region

        .EXAMPLE
            Get-eduSTARMCSchool
    #>
	
	[xml]$Request = Invoke-RestMethod -Uri https://apps.edustar.vic.edu.au/edustarmc/api/MC/GetAllSchools -Method Get -WebSession $eduSTARMCSession -ContentType "application/xml"
	
	$Result = @()
	
	ForEach ($obj in $Request.ArrayOfSchool.School)
	{
		$item = New-Object PSObject -Property @{
			SchoolId = $obj.SchoolId
			SchoolName = $obj.SchoolName
		}
		
		$Result += $item
	}
	
	$ResultArray = @($Result)
	
	return $ResultArray
	
}

function Select-eduSTARMCSchool
{
	
	if ($null -eq $script:eduSTARMCSchools)
	{
		$script:eduSTARMCSchools = Get-eduSTARMCSchool
	}
	
	$SelectedSchool = $null
	
	if ($eduSTARMCConnection.Schools -eq 1)
	{
		$SelectedSchool = $eduSTARMCSchools[0]
	}
	else
	{
		$SelectedSchool = $eduSTARMCSchools | Out-GridView -Title 'Select a school' -PassThru -ErrorAction Stop
	}
	
	return $SelectedSchool.SchoolId
}

#endregion

#region Groups

function Get-eduSTARMCGroup
{
    <#
        .SYNOPSIS
            Returns all groups for the specified school

        .OUTPUTS
            Returns Username, FirstName, LastName and Enabled (True/False)

        .EXAMPLE
            Get-eduSTARMCGroup

            This returns all groups from the selected school

        .EXAMPLE

            Get-eduSTARMCGroup -SchoolNumber 1234

            This returns all groups from a specific school

        .EXAMPLE

            Get-eduSTARMCGroup -SchoolNumber 1234 -Identity "1234-ls-Local Technician"

            This returns a specific group from a specific school
    #>
	
	param (
		[int]$SchoolNumber = (Select-eduSTARMCSchool),
		[string]$Identity,
		[int]$Timeout = 60
	)
	
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/GetSchoolGroups/{0}" -f $SchoolNumber)
	
	$Request = Invoke-RestMethod -Uri $Uri -Method Get -WebSession $eduSTARMCSession -ContentType "application/json"
	
	$Result = ($Request._centralGroups + $Request._localGroups) | Select-Object @{ label = "GroupName"; expression = { $_._groupName } }, @{ label = "DistinguishedName"; expression = { $_._dn } }
	
	# Return the group based on name
	if ([string]::IsNullOrEmpty($Identity))
	{
		Write-Output $Result
	}
	else
	{
		Write-Output $Result.Where({ $_.GroupName -eq $Identity })
	}
}

function Get-eduSTARMCGroupMember
{
    <#
        .SYNOPSIS
            Returns all groups for the specified school

        .OUTPUTS
            Returns FirstName, LastName, Identity, Common Name, DistinguishedName, CanDelete, EmailAddress

        .EXAMPLE
            Get-eduSTARMCGroupMember -SchoolNumber 1234

            Get-eduSTARMCGroupMember -SchoolNumber 1234 -GroupName "1234-gs-Local Technician"
    #>
	
	param (
		[int]$SchoolNumber = (Select-eduSTARMCSchool),
		[Parameter(Mandatory = $true)]
		[string]$GroupName,
		[int]$Timeout = 60
	)
	
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	$GroupDN = (Get-eduSTARMCGroup -Identity $GroupName -SchoolNumber $SchoolNumber).DistinguishedName
	
	$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/GetGroupMembers?schoolId={0}&groupDn={1}&groupName={2}" -f $SchoolNumber, $GroupDN, $GroupName)
	
	$Request = Invoke-RestMethod -Uri $Uri -Method Get -WebSession $eduSTARMCSession -ContentType "application/json"
	
	$Result = $Request
	
	Write-Output $Result
}

#endregion

#region Accounts

function Get-eduSTARMCServiceAccount
{
    <#
        .SYNOPSIS
            Returns all schools service accounts for the specified school

        .OUTPUTS
            Returns Username, FirstName, LastName and Enabled (True/False)

        .EXAMPLE
            Get-eduSTARMCServiceAccount -SchoolNumber 1234
    #>
	
	param (
		[int]$SchoolNumber = (Select-eduSTARMCSchool),
		[int]$Timeout = 60
	)
	
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/GetSchoolServiceAccounts/{0}" -f $SchoolNumber)
	
	$Request = Invoke-RestMethod -Uri $Uri -Method Get -WebSession $eduSTARMCSession -ContentType "application/json" -TimeoutSec $Timeout
	
	$Result = $Request | Select @{ label = "Username"; expression = { $_._login } }, @{ label = "FirstName"; expression = { $_._firstName } }, @{ label = "LastName"; expression = { $_._lastName } }, @{ label = "Enabled"; expression = { !$_._disabled } }
	
	Write-Output $Result
}

function Get-eduPassAccount
{
    <#
        .SYNOPSIS
            Retrieves information about a student account
            Downloads the entire list of students to cache on first run

        .PARAMETER Force
            Clears the student cache before retrieving students

        PARAMETER Identity
            Return a specific user

        .PARAMETER Timeout
            Optional parameter to set timeout of the request (default is 60 seconds)
            
        .EXAMPLE

            Get-eduPassAccount

            Gets accounts from eduSTAR MC, utilising cache if applicable

        .EXAMPLE

            Get-eduPassAccount | Out-GridView

            Gets accounts from eduSTAR MC and outputs to a gridview

        .EXAMPLE

            Get-eduPassAccount -Identity jsmith

            Gets specified account from eduSTAR MC

        .EXAMPLE

            Get-eduPassAccount -Force

            Gets accounts from eduSTAR MC, bypassing cache
    #>
	
	param (
		[string]$Identity,
		[int]$SchoolNumber = (Select-eduSTARMCSchool),
		[switch]$Force,
		[int]$Timeout = 60
	)
	
	$CacheRootPath = "$env:TEMP\eduSTARMCAdministration"
	$StudentCache = ("{0}\{1}-Students.xml" -f $CacheRootPath, $SchoolNumber)
	
	$eduPassAccounts = @()
	
	# Check if the cache folder exists before trying to save the XML to temp folder
	if (-not (Test-Path -Path $CacheRootPath))
	{
		New-Item -Path $CacheRootPath -ItemType Directory -Force | Out-Null
	}
	
	# Check if school number is valid
	#Test-eduSTARMCSchoolNumber $SchoolNumber
	
	# If the array is not loaded into memory, attempt to load it from cache (temp folder)
	if (-not $Force.IsPresent)
	{
		try
		{
			$Cache = Get-Content -Path $StudentCache -ErrorAction SilentlyContinue
			$eduPassAccounts = ([xml]$Cache).ArrayOfStudent
		}
		catch
		{
			Write-Output 'Error reading from cache'
			$eduPassAccounts = $null
		}
	}
	
	if ($null -eq $eduPassAccounts -or $Force.IsPresent)
	{
		# If the array is still empty (cache doesn't exist or failed to read), retrieve new data from the eduSTAR MC
		
		Write-Progress -Activity 'Please wait while student data is retrieved from the eduSTAR Management Console...' -PercentComplete 10
		
		$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/GetStudents/{0}/FULL" -f $SchoolNumber)
		
		[xml]$Request = Invoke-RestMethod -Uri $Uri -Method Get -WebSession $eduSTARMCSession -ContentType "application/xml" -TimeoutSec $Timeout
		
		
		$eduPassAccounts = ([xml]$Request).ArrayOfStudent
		
		# Save the XML file to temp folder to act as cache
		$Request.Save($StudentCache)
		
		# Hide the progress bar
		Write-Progress -Activity 'Retrieving all students' -Completed
	}
	
	
	# Return the student based on login/username
	if ([string]::IsNullOrEmpty($Identity))
	{
		Add-eduPassAccountsAlias($eduPassAccounts.Student)
	}
	else
	{
		Add-eduPassAccountsAlias($eduPassAccounts.Student) | Where-Object { $_.login -eq $Identity }
	}
}

function Add-eduPassAccountsAlias
{
	# Strips the leading underscore from 
	param (
		[array]$eduPassAccounts
	)
	
	($eduPassAccounts | Get-Member -MemberType Property).Name | ForEach-Object {
		$eduPassAccounts | Add-Member -MemberType AliasProperty -Name $_.Trim("_") -Value $_
	}
	
	return $eduPassAccounts
}

function Set-eduPassAccountPassword
{
    <#
        .SYNOPSIS
            Sets the password of a student


        .PARAMETER DN
            Distinguished Name of the student
            

        .EXAMPLE

            Set-eduPassAccountPassword

            This will prompt to select a user and resets to a random password
        

        .EXAMPLE

            Set-eduPassAccountPassword -Identity jsmith -Password "Mypassword123!"

            This will set the password for John Smith using the specified password


        .EXAMPLE

            Set-eduPassAccountPassword -Password "Str0ngpa$$word!"

            This will prompt to select a user and resets to the specified password
    #>
	
	param (
		[string]$Identity,
		[int]$SchoolNumber = (Select-eduSTARMCSchool),
		[ValidateNotNullOrEmpty()]
		[string]$Password
	)
	
	$Account = $null
	
	# Check if school number is valid
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	if ($PSBoundParameters.ContainsKey('Identity'))
	{
		$Account = Get-eduPassAccount -Identity $Identity -SchoolNumber $SchoolNumber
	}
	else
	{
		$Account = Get-eduPassAccount -SchoolNumber $SchoolNumber | Out-GridView -PassThru -ErrorAction Stop
	}
	
	if ([string]::IsNullOrEmpty($Account.Dn))
	{
		return "No user was selected, exiting..."
	}
	
	# If a password wasn't specified, get a random one from dinopass
	if ([string]::IsNullOrEmpty($Password))
	{
		$RandomPassword = Invoke-RestMethod -Uri 'https://www.dinopass.com/password/simple' -Method Get
		
		$Password = ("{0}!" -f (Get-Culture).TextInfo.ToTitleCase($RandomPassword))
	}
	
	Add-Type -AssemblyName System.Web
	
	if ($null -ne $eduSTARMCSession)
	{
		
		$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/ResetStudentPwd")
		
		$Parameters = @{
			dn = $Account.dn
			newPass = $Password
			schoolId = $SchoolNumber
		} | ConvertTo-Json
		
		try
		{
			$Request = Invoke-RestMethod -Uri $Uri -WebSession $eduSTARMCSession -Method Post -Body $Parameters -ContentType "application/json"
			
			$Result = New-Object PSObject -Property ([ordered]@{
					Name	 = ("{0} {1}" -f $Account.firstName, $Account.lastName)
					Username = $Account.login
					Password = $Password
				})
			
			Write-Output $Result | Format-List
			
		}
		catch
		{
			Write-Error -Message ("Unable to set password for '{0}'. Detail: {1}" -f $Account.login, $_)
		}
		
	}
	else
	{
		throw 'You must open a session to the eduSTAR Management Console before using this function'
	}
}

#endregion

#region Cloud

function Set-eduPassCloudServiceStatus
{
    <#
        .SYNOPSIS
            Retrieves information about a student account
            Downloads the entire list of students to cache on first run

        PARAMETER Identity
            Specify usernames of user(s)

        .PARAMETER Timeout
            Optional parameter to set timeout of the request (default is 60 seconds)
            
        .EXAMPLE
            Set-eduPassCloudServiceStatus -Identity jsmith -Service google -AccountType student -Status Enabled

            Set-eduPassCloudServiceStatus -Identity jdoe, hpotter -SchoolNumber 1234 -Service google -AccountType staff -Status Disabled

    #>
	
	param (
		[ValidateNotNullOrEmpty()]
		[string[]]$Identity,
		$SchoolNumber,
		[ValidateNotNullOrEmpty()]
		[Parameter()]
		[ValidateSet('staff', 'student', 'serviceaccount')]
		[string]$AccountType,
		[ValidateNotNullOrEmpty()]
		[Parameter()]
		[ValidateSet('o365', 'intune', 'google', 'yammer', 'lynda', 'stile', 'webex')]
		[string]$Service,
		[ValidateNotNullOrEmpty()]
		[Parameter()]
		[ValidateSet('Enabled', 'Disabled')]
		[string]$Status,
		[int]$Timeout = 60
	)
	
	[array]$DNarray = @()
	
	
	$Accounts = @()
	ForEach ($i in $Identity)
	{
		$Accounts += Get-eduPassAccount -Identity $i -SchoolNumber $SchoolNumber
		
	}
	
	ForEach ($Account in $Accounts)
	{
		$DNarray += $Account.Dn
	}
	
	
	$WebRequestBody = @{
		_accountType = $AccountType
		_dns		 = $DNarray
		_schoolId    = $SchoolNumber
		_property    = $Service
	} | ConvertTo-Json
	
	$endpoint = $null
	
	if ($Status -eq 'Enabled')
	{
		$endpoint = "SetO365"
	}
	else
	{
		$endpoint = "UnsetO365"
	}
	
	$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/{0}" -f $endpoint)
	
	$Result = $null
	
	try
	{
		$Request = Invoke-RestMethod -Uri $Uri -WebSession $eduSTARMCSession -Body $WebRequestBody -Method Post -TimeoutSec $Timeout -ContentType "application/json"
		
		ForEach ($Account in $Accounts)
		{
			$row = New-Object PSObject -Property @{
				Username = $Account.login
				Name	 = ("{0} {1}" -f $Account.lastName, $Account.firstName)
				Service  = $Service
				Status   = $Status
			}
			
			$Result += $row
		}
	}
	catch
	{
		$Result = "Failed to update cloud service."
	}
	
	Write-Output $Result
	
}

#endregion

#region Distribution Lists

function New-eduSTARMCDistributionList
{
    <#
        .SYNOPSIS
           


        .PARAMETER DN
           
            

        .EXAMPLE

        
    #>
	
	param (
		[Parameter(Mandatory = $true)]
		[string]$Name,
		[int]$SchoolNumber = (Select-eduSTARMCSchool)
	)
	
	# Check if school number is valid
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	
	if ($null -ne $eduSTARMCSession)
	{
		
		$DistributionListName = ("{0}-dl-{1}" -f $SchoolNumber, $Name)
		$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/CreateDL?schoolId={0}&dlName={1}" -f $SchoolNumber, $DistributionListName)
		
		try
		{
			$Request = Invoke-RestMethod -Uri $Uri -WebSession $eduSTARMCSession -Method Post
			
			$Result = New-Object PSObject -Property ([ordered]@{
					DistributionListName = $DistributionListName
					Action			     = "Created"
				})
			
			Write-Output $Result
			
		}
		catch
		{
			Write-Error -Message ("Unable to create new distribution list '{0}'. Detail: {1}" -f $DistributionListName, $_)
		}
		
	}
	else
	{
		throw 'You must open a session to the eduSTAR Management Console before using this function'
	}
}

function Remove-eduSTARMCDistributionList
{
    <#
        .SYNOPSIS
           


        .PARAMETER DN
           
            

        .EXAMPLE

        
    #>
	
	param (
		[Parameter(Mandatory = $true)]
		[string]$Name,
		[int]$SchoolNumber = (Select-eduSTARMCSchool)
	)
	
	# Check if school number is valid
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	
	if ($null -ne $eduSTARMCSession)
	{
		
		$DistributionListName = ("{0}-dl-{1}" -f $SchoolNumber, $Name)
		$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/DeleteDL?schoolId={0}&dlName={1}" -f $SchoolNumber, $DistributionListName)
		
		try
		{
			$Request = Invoke-RestMethod -Uri $Uri -WebSession $eduSTARMCSession -Method Post
			
			$Result = New-Object PSObject -Property ([ordered]@{
					DistributionListName = $DistributionListName
					Action			     = "Removed"
				})
			
			Write-Output $Result
			
		}
		catch
		{
			Write-Error -Message ("Unable to create new distribution list '{0}'. Detail: {1}" -f $DistributionListName, $_)
		}
		
	}
	else
	{
		throw 'You must open a session to the eduSTAR Management Console before using this function'
	}
}

function Get-eduSTARMCDistributionList
{
    <#
        .SYNOPSIS
           

        .OUTPUTS
            

        .EXAMPLE
            

        
    #>
	
	param (
		[string]$Name,
		[int]$SchoolNumber = (Select-eduSTARMCSchool),
		[int]$Timeout = 60
	)
	
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/GetSchoolDL/{0}" -f $SchoolNumber)
	
	$Request = Invoke-RestMethod -Uri $Uri -Method Get -WebSession $eduSTARMCSession -ContentType "application/json"
	
	$Result = $Request | Select-Object -Property @{ Name = 'Name'; Expression = { $_._groupName } }
	
	# Return the group based on name
	if ([string]::IsNullOrEmpty($Identity))
	{
		Write-Output $Result
	}
	else
	{
		Write-Output $Result.Where({ $_._groupName -eq $Identity })
	}
}

function Get-eduSTARMCDistributionListMember
{
    <#
        .SYNOPSIS
            Returns members from the specified distribution group

        .OUTPUTS
            Returns Username, Name and CanDelete (boolean)

        .EXAMPLE
            Get-eduSTARMCDistributionListMember -DistributionListName 'XXXX-name'

    #>
	
	param (
		[int]$SchoolNumber = (Select-eduSTARMCSchool),
		[Parameter(Mandatory = $true)]
		[string]$DistributionListName
	)
	
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/GetDLMembers/{0}/{1}" -f $SchoolNumber, $DistributionListName)
	
	$Request = Invoke-RestMethod -Uri $Uri -Method Get -WebSession $eduSTARMCSession -ContentType "application/json"
	
	$Result = $Request | Select-Object -Property @(
		@{ Name = 'Username'; Expression = { $_._login } }
		@{ Name = 'Name'; Expression = { $_._dn } }
		@{ Name = 'CanDelete'; Expression = { $_._canDelete } }
	)
	
	Write-Output $Result
}

function Add-eduSTARMCDistributionListMember
{
    <#
        .SYNOPSIS
            


        .PARAMETER DN
            
            

        .EXAMPLE

     
    #>
	
	param (
		[ValidateNotNullOrEmpty()]
		[string]$Identity,
		[int]$SchoolNumber = (Select-eduSTARMCSchool),
		[ValidateNotNullOrEmpty()]
		[string]$DistributionListName
	)
	
	$Account = $null
	
	# Check if school number is valid
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	
	$Account = Get-eduPassAccount -Identity $Identity -SchoolNumber $SchoolNumber
	
	
	if ([string]::IsNullOrEmpty($Account.DistinguishedName))
	{
		return "No user was selected, exiting..."
	}
	
	if ($null -ne $eduSTARMCSession)
	{
		
		$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/AddDLMember?schoolId={0}&dlName={1}&memberId={2}&memberDisplayName={3}" -f $SchoolNumber, $DistributionListName, $Account.Username, $Account.DistinguishedName)
		
		try
		{
			$Request = Invoke-RestMethod -Uri $Uri -WebSession $eduSTARMCSession -Method Post
			
			$Result = New-Object PSObject -Property ([ordered]@{
					Name	 = ("{0} {1}" -f $Account.Givenname, $Account.Surname)
					Username = $Account.Username
					DistributionListName = $DistributionListName
					Action   = "Added"
				})
			
			Write-Output $Result | Format-List
			
		}
		catch
		{
			Write-Error -Message ("Unable to remove user '{0}' from '{1}'. Detail: {1}" -f $Account.Username, $DistributionListName, $_)
		}
	}
	else
	{
		throw 'You must open a session to the eduSTAR Management Console before using this function'
	}
}

function Remove-eduSTARMCDistributionListMember
{
    <#
        .SYNOPSIS
            


        .PARAMETER DN
            
            

        .EXAMPLE

     
    #>
	
	param (
		[ValidateNotNullOrEmpty()]
		[string]$Identity,
		[int]$SchoolNumber = (Select-eduSTARMCSchool),
		[ValidateNotNullOrEmpty()]
		[string]$DistributionListName
	)
	
	$Account = $null
	
	# Check if school number is valid
	Test-eduSTARMCSchoolNumber $SchoolNumber
	
	
	$Account = Get-eduPassAccount -Identity $Identity -SchoolNumber $SchoolNumber
	
	if ([string]::IsNullOrEmpty($Account.DistinguishedName))
	{
		return "No user was selected, exiting..."
	}
	
	if ($null -ne $eduSTARMCSession)
	{
		
		$Uri = ("https://apps.edustar.vic.edu.au/edustarmc/api/MC/RemoveDLMember?schoolId={0}&dlName={1}&memberId={2}" -f $SchoolNumber, $DistributionListName, $Account.Username)
		
		try
		{
			$Request = Invoke-RestMethod -Uri $Uri -WebSession $eduSTARMCSession -Method Post
			
			$Result = New-Object PSObject -Property ([ordered]@{
					Name	 = ("{0} {1}" -f $Account.Givenname, $Account.Surname)
					Username = $Account.Username
					DistributionListName = $DistributionListName
					Action   = "Removed"
				})
			
			Write-Output $Result | Format-List
			
		}
		catch
		{
			Write-Error -Message ("Unable to add user '{0}' to '{1}'. Detail: {1}" -f $Account.Username, $DistributionListName, $_)
		}
	}
	else
	{
		throw 'You must open a session to the eduSTAR Management Console before using this function'
	}
}

#endregion

#region Helpers

function Format-eduSTARMCDateTime
{
	param (
		[Parameter(Mandatory = $true)]
		$DateTime
	)
	
	if ($DateTime -is [System.String])
	{
		$result = New-Object DateTime
		$Format = 'yyyy-MM-ddTHH:mm:ss'
		
		$DateSplit = $DateTime.Split('.')[0]
		
		$ValidDate = [datetime]::TryParseExact(
			$DateSplit,
			$Format,
			[System.Globalization.CultureInfo]::InvariantCulture,
			[System.Globalization.DateTimeStyles]::None,
			[ref]$result
		)
		
		if ($ValidDate)
		{
			Write-Output $result
		}
	}
	else
	{
		Write-Output $null
	}
}

function Test-eduSTARMCSchoolNumber
{
	param (
		[Parameter(Mandatory = $true)]
		[int]$SchoolNumber
	)
	
	if ($SchoolNumber -notmatch "[0-9][0-9][0-9][0-9]")
	{
		throw 'Unable to retrieve data without selecting a school'
	}
}

#endregionf

#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory



