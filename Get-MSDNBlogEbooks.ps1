#==================================================================================================#
#
# Author:				Collin Chaffin
# Last Modified:		07-12-2015 01:00 PM
# Filename:				Get-MSDNBlogEbooks.ps1
#
#
# Changelog:
#
#	v 1.0.0.1	:	07-16-2015	:	Initial release
#
# Notes:
#
#	This script automates downloading all free ebooks posted to the
#	following Microsoft MSDN blog:  http://bit.ly/FreeMSDNbooks
#
#	It Optionally will create a local zip file for archiving containing
#	all the downloaded ebook files.
#
#	Example:
#		Get-MSDNBlogEbooks.ps1 -SavePath "C:\Data\eBooks" -CreateZip -ZipFile "C:\Temp\eBooks.zip"
#
#
#
#    THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
#    ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
#    THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
#    PARTICULAR PURPOSE AND NONINFRINGEMENT.
#
#    IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
#    FOR ANY SPECIAL, INDIRECT OR CONSEQUENTIAL DAMAGES OR ANY
#    DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
#    WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS
#    ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE
#    OF THIS CODE OR INFORMATION.
#
#
#
#	**NOTE -- NO EDITING SHOULD BE REQUIRED IF THE REQUIRED PARAMS ARE PROVIDED ON COMMAND LINE**
#	**NOTE -- AT TIME OF AUTHORING, SEVERAL OF THE PROVIDED EBOOK URLS WERE NOT REACHABLE**
#	**NOTE -- THIS DOWNLOADS OVER 1.2 GIGABYTES OF EBOOKS.  EVEN ZIPPING IS OVER A GIG!!!!!!!!
#
#==================================================================================================#

[CmdletBinding(DefaultParameterSetName = 'NoZip')]
param
(
	[Parameter(ParameterSetName = 'NoZip',
			   Mandatory = $true)]
	[Parameter(ParameterSetName = 'Zip',
			   Mandatory = $true)]
	[string]
	$SavePath,
	[Parameter(ParameterSetName = 'Zip',
			   Mandatory = $true)]
	[switch]
	$CreateZip,
	[Parameter(ParameterSetName = 'Zip',
			   Mandatory = $true)]
	[string]
	$ZipFile
)


#region Globals
#########################################################################
# 							Global Variables							#
#########################################################################

# Paths
$webEbookList = "http://bit.ly/freeMSDNbooklist"
$localBookList = "$env:temp\MSFTEbooks2.txt"
$invocationPath = $([System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition) + "\")
#If the zip file target passed in does not have a full path or is invalid
#save to the current folder from which script is executing
switch ($PsCmdlet.ParameterSetName)
{
	'NoZip' {
		$ZipFile = $null
		break
	}
	'Zip' {
		if ((!($([System.IO.Path]::GetDirectoryName($zipFile)))) -or (!(Test-Path -Path ($([System.IO.Path]::GetDirectoryName($zipFile)))))) { $ZipFile = "$($invocationPath)$($([System.IO.Path]::GetFileName($ZipFile)))" }
		break
	}
}


#########################################################################
#endregion



#region Functions

#########################################################################
# 								Functions								#
#########################################################################


function Get-ebookListtoCSV
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.String]
		$EbookListURL,
		[Parameter(Mandatory = $true)]
		[System.String]
		$LocalTempCSV
	)
	try
	{
		#Retrieve ebook list
		Invoke-WebRequest -Uri $EbookListURL -OutFile $LocalTempCSV
		Unblock-File $LocalTempCSV
	}
	catch
	{
		Throw $("ERROR OCCURRED WHILE RETRIEVING EBOOK LIST " + $_.Exception.Message)
	}
}

function Get-eBookFiles
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.String]
		$SavePath,
		[Parameter(Mandatory = $true)]
		[System.String]
		$InputCSV
	)
	try
	{
		if (!(Test-Path -Path $SavePath)) { New-Item -ItemType Directory -Force -Path $SavePath | Out-Null }
		#Set-Location -Path $SavePath
		$eBooks = (Import-Csv -Path $InputCSV -Header 'ebook' | select -skip 1)
		foreach ($eBook in $eBooks)
		{
			#Download the ebooks one at a time using my function
			Get-HTTPFile -FileURL "$($eBook.ebook)" -FileOut "$($SavePath)" -ErrorAction 'Continue'
		}
	}
	catch
	{
		Throw $("ERROR OCCURRED WHILE DOWNLOADING EBOOKS " + $_.Exception.Message)
	}
}

function New-ZipArchive
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.String]
		$SourceDirectory,
		[Parameter(Mandatory = $true)]
		[System.String]
		$ZipArchive
	)
	try
	{
		#Load assembly. Please don't tell me to use add-type - I am not the only dev been bitten by using add-type
		#until they finally remove depreciated support of LoadWithPartialName, I'll use it.
		[Reflection.Assembly]::LoadWithPartialName("System.IO.Compression.FileSystem") | Out-Null
		
		#Avoid the ugly red error if the zip already exists
		If (Test-path $ZipArchive) { Remove-item $ZipArchive -Force | Out-Null }
		
		Write-Host "`nCreating ZIP archive - please wait....`n" -ForegroundColor 'Yellow'
		
		#Zip up the new folder we just filled with ebooks - dotnet sucks ass on zipping with any progress reporting....so nope - no progress hence the write-hosts
		#so the huge lockup while zipping over a gig takes place :)
		[io.compression.zipfile]::CreateFromDirectory($SourceDirectory, $ZipArchive)
		
		Write-Host "`nZIP archive created successfully!`n" -ForegroundColor 'Yellow'
	}
	catch
	{
		Throw $("ERROR OCCURRED WHILE LOADING REQUIRED DOTNET ASSEMBLIES AND CREATING ZIP ARCHIVE " + $_.Exception.Message)
	}	
}

function Get-HTTPFileName
{
	[CmdletBinding()]
	[OutputType([string])]
	param
	(
		[Parameter(Mandatory = $true)]
		[System.String]
		$FileURL = (Read-Host "URL of file")
	)	
	BEGIN
	{
	}
	PROCESS
	{
		try
		{
			#We were just born, so everything is okay
			$errCode = 0
			
			#Send the request
			$urlRequest = [System.Net.HttpWebRequest]::Create($FileURL)
			$urlRequest.AllowAutoRedirect = $true
			$urlRequest.Set_Timeout(15000)
			
			#Get the response
			$urlResponse = $urlRequest.GetResponse()
			
			#If we were provided local filename without full path, write it to current folder
			if ($FileOut -and !(Split-Path $FileOut))
			{
				$FileOut = Join-Path (Get-Location -PSProvider "FileSystem") $FileOut
			}
			#If we weren't provided a local filename to save as, or if we were provided
			#only a folder destination and want auto filename, save the path to prepend
			#and continue with retrieving filename
			elseif ((!($FileOut)) -or (Test-Path -PathType "Container" $FileOut))
			{
				#Store the save path
				$savePath = $FileOut
				
				#Parse the headers for Content-Disposition, then first check if "filename=" is properly provided
				[string]$fileName = ([regex]'(?i)filename=(.*)$').Match($urlResponse.Headers["Content-Disposition"]).Groups[1].Value
				
				#Crap - it's really not always that simple.  Let's try a second method to get the filename.
				if (!$fileName)
				{
					#If "filename" not provided in header with Content-Disposition, we will attempt to pull from
					#the ResponseUri segments property
					$fileName = $urlResponse.ResponseUri.Segments[-1]
				}
				
				#If we tried our two methods above and found the filename, perform a trim and verify/add extension
				if ($fileName)
				{
					#Cleanup the filename and verify and if required add extension.  If we still failed we
					#are done - exit with an error. User must rerun and force a local filename
					
					#Trim any slashes
					$fileName = $fileName.Trim("\/")
					
					#Cast what we think is our filename as a IO.FileInfo type to then verify whether we
					#believe it has a proper file extension.  If not, parse the perceived Content-Type and
					#append our own extension of that perceived mime type
					if (!([IO.FileInfo]$fileName).Extension)
					{
						#$fileName = $fileName + "." + $urlResponse.ContentType.Split(";")[0].Split("/")[1]
						$fileName = $fileName + "." + ([System.Web.MimeMapping]::GetMimeMapping($fileName)).Split(";")[0].Split("/")[1]
					}
				}
				#We failed so return failure code to be handled below
				else
				{
					$errCode = 1
				}
			}
			
		}
		catch
		{
			#Throw our general exception and die a quick death
			$errCode = 1
			#Throw $("ERROR OCCURRED WHILE DETERMINING HTTP FILENAME " + $_.Exception.Message)
			Continue
		}
		finally
		{
			#If we have any status code other than zero, die a quick death
			If ($errCode -ne 0)
			{
				Write-Host "ERROR OCCURRED WHILE FETCHING URL: $(Truncate -strInput $FileURL -maxLength 14)" -ForegroundColor 'Red'
				#[Environment]::Exit($errCode)
			}
			
			#Close handles and flush/dispose			
			if ($urlResponse) { $urlResponse.Dispose() }
		}
	}
	END
	{
		return $fileName;
	}
}


function Get-HTTPFile
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[System.String]
		$FileURL = (Read-Host "URL of file"),
		[Parameter(Mandatory = $false)]
		[System.String]
		$FileOut
	)	
	BEGIN
	{
		$_savePath = $null
		$_fileOut = $FileOut
	}
	PROCESS
	{
		try
		{
			#We were just born, so everything is okay
			$errCode = 0
			
			#If we were provided local filename without full path, write it to current folder
			if ($_fileOut -and (!(Split-Path $_fileOut)))
			{
				$_fileOut = Join-Path (Get-Location -PSProvider "FileSystem") $_fileOut
			}
			#If we weren't provided a local filename to save as, or if we were provided
			#only a folder destination and want auto filename, save the path to prepend
			#and continue with retrieving filename
			elseif ((!($_fileOut)) -or (Test-Path -PathType "Container" $_fileOut))
			{
				#Store the save path accounting for root folders
				if (Test-Path -PathType "Container" $_fileOut)
				{
					if ($_fileOut.Substring($_fileOut.Length - 1, 1) -eq '\')
					{
						$_savePath = $_fileOut
					}
					else
					{
						$_savePath = "$($_fileOut)\"
					}
				}
				
				#Get the filename from the webserver
				$_fileOut = (Get-HTTPFileName -FileURL "$($FileURL)")
				
				#We got the filename, let's decide where to store it
				$_fileOut = "$($_savePath)$($_fileOut)"
				
			}
		}
		catch
		{
			#Throw our general exception and die a quick death
			$errCode = 1
			#Change to continue because a terminating error here will halt multiple downloads and we don't want that
			Continue
		}
		finally
		{
			#If we have any status code other than zero, die a quick death
			If ($errCode -ne 0)
			{
				Write-Host "ERROR OCCURRED WHILE DETERMINING HTTP FILENAME" -ForegroundColor 'Red'
			}
		}
		
		
		#Download the file
		try
		{
			#We were just born, so everything is okay
			$errCode = 0
			
			#Send the request
			$urlRequest = [System.Net.HttpWebRequest]::Create($FileURL)
			$urlRequest.AllowAutoRedirect = $true
			$urlRequest.Set_Timeout(15000)
			
			#Get the response
			$urlResponse = $urlRequest.GetResponse()
			
			#Get the file size but if it is less than 1k set it to 1k to avoid math errors
			if ($($urlResponse.get_ContentLength()) -lt 1024)
			{
				$totalLength = [System.Math]::Floor(1024/1024)
			}
			else
			{
				$totalLength = [System.Math]::Floor($urlResponse.get_ContentLength()/1024)
			}
			
			#Get the file stream
			$urlStream = $urlResponse.GetResponseStream()
			
			#Open the file for writing
			$writeFile = New-Object -TypeName System.IO.FileStream -ArgumentList $_fileOut, Create
			$byteBuffer = New-Object byte[] 10KB
			$byteCount = $urlStream.Read($byteBuffer, 0, $byteBuffer.length)
			$downloadedBytes = $byteCount
			
			#Loop through every byte and display progress
			while ($byteCount -gt 0)
			{
				#Use csharp method to write to console - better same-line progress and less obtrusive than Write-Progress
				[System.Console]::CursorLeft = 0
				[System.Console]::BackGroundColor = 'Black'
				[System.Console]::ForeGroundColor = 'Yellow'				
				[System.Console]::Write("Downloading {0}:  {1}K of {2}K", $(Truncate-WithExtension -strInput $_fileOut -strExtension $($_fileOut.Substring($_fileOut.Length - 3, 3)) -maxLength 48), [System.Math]::Floor($downloadedBytes/1024), $totalLength)
				
				#Write the binary data in buffered chunks of 10k
				$writeFile.Write($byteBuffer, 0, $byteCount)
				$byteCount = $urlStream.Read($byteBuffer, 0, $byteBuffer.length)
				$downloadedBytes = $downloadedBytes + $byteCount
			}
			[System.Console]::Write("`n")
			[System.Console]::ResetColor()
						
		}
		catch
		{
			#Throw our general exception and die a quick death
			$errCode = 1
			#Throw $("ERROR OCCURRED WHILE DOWNLOADING FILE " + $_.Exception.Message)
			Continue
		}
		finally
		{
			#If we have any status code other than zero, die a quick death
			If ($errCode -ne 0)
			{
				Write-Host "ERROR OCCURRED WHILE DOWNLOADING FILE $($_fileOut)" -ForegroundColor 'Red'
				#Read-Host
				#[Environment]::Exit($errCode)
			}
			
			#Close handles and flush/dispose
			if ($writeFile) { $writeFile.Flush() }
			if ($writeFile) { $writeFile.Close() }
			if ($writeFile) { $writeFile.Dispose() }
			if ($urlStream) { $urlStream.Dispose() }
			if ($urlResponse) { $urlResponse.Dispose() }			
		}
	}
	END
	{
	}
}

function Truncate
{
	[CmdletBinding()]
	[OutputType([System.String])]
	param
	(
		[Parameter(Mandatory = $true)]
		[System.String]
		$strInput,
		[Parameter(Mandatory = $true)]
		[System.Int32]
		$maxLength
	)
	if ($strInput.Length -gt $maxLength)
	{
		return $($strInput.Substring(0, $maxLength) + "..");
	}
	else
	{
		return $strInput;
	}
}

function Truncate-WithExtension
{
	[CmdletBinding()]
	[OutputType([System.String])]
	param
	(
		[Parameter(Mandatory = $true)]
		[System.String]
		$strInput,
		[Parameter(Mandatory = $true)]
		[System.String]
		$strExtension,
		[Parameter(Mandatory = $true)]
		[System.Int32]
		$maxLength
	)
	if ($strInput.Length -gt $maxLength)
	{
		return $($strInput.Substring(0, ($maxLength - 7)) + "..." + $strExtension);
	}
	else
	{
		return $strInput;
	}
}


#########################################################################
#endregion



#region Program Execution
#########################################################################
# 						MAIN PROGRAM EXECUTION							#
#########################################################################

#Clear host
	Clear-Host

#Retrieve ebook list
	Get-ebookListtoCSV -EbookListURL $webEbookList -LocalTempCSV $localBookList


#Move to destination path and get the ebooks
	Get-eBookFiles -SavePath $SavePath -InputCSV $localBookList


#Set up optional zip archive
	switch ($PsCmdlet.ParameterSetName)
	{
		'NoZip' {
			Write-Host "`n`nComplete!`n`n" -ForegroundColor 'Green'
			break
		}
		'Zip' {
			New-ZipArchive -SourceDirectory $SavePath -ZipArchive $zipFile
			Write-Host "`n`nComplete!`n`n" -ForegroundColor 'Green'
			break
		}
	}


#########################################################################
#endregion