<#
    Function: Install and configure SQL Server
    Updated on： 03/10/2021
    Created by: Shane
#>
function GetInformation {
    param (
        [validateset(` 
        'SQL_2019_Express_X64',
        'SQL_2017_Express_X64',
        'SQL_2016_Express_X64',
        'SQL_2014_Express_X64',
        'SQL_2012_Exp_SP1_x64',
        'SQL_2008_R2_Exp_SP2_x64',
        'SQL_2019_Web_X64',
        'SQL_2017_Web_X64',
        'SQL_2016_Web_X64',
        'SQL_2014_Web_x64',
        'SQL_2012_Web_SP1_x64',
        'SQL_2008_R2_Web_SP2',
        'SQL_2019_STD_X64',
        'SQL_2017_STD_X64',
        'SQL_2016_STD_X64',
        'SQL_2014_STD_x64',
        'SQL_2012_STD_SP1_x64'`
        )]
        [string]$SQLServerVersion,
        [string]$FtpPassword,
        [string]$SaPassword
    )
    [pscustomobject]@{
        SQLServerVersion = $SQLServerVersion
        FtpPassword = $FtpPassword
        SaPassword = [string]$SaPassword
    }
}
$InstallInfo = Invoke-Expression (Show-Command GetInformation -PassThru)
if(!$?) {
    Write-Host "Click cancel. End of this installation!" -ForegroundColor Red
    break;
}

#Get installation information
$SQLVersion = $InstallInfo.SQLServerVersion
$FTPPassword = $InstallInfo.FtpPassword
$saPwd = $InstallInfo.SaPassword

#Check inputed information
if (!($SQLVersion -and $FTPPassword -and $saPwd)) {
    Write-Host "You didn't enter the full information, please rerun the script and enter the necessary information!" -ForegroundColor Red
    break
}

#SQL File name on FTP
$SQLFile = $SQLVersion + ".zip"

#Get FTP file size
function Get-FTPFilesize ([String] $FTPUser, [String] $FTPPassword, [String] $FTPLink) 
{ 
	$FTPRequest = [System.Net.FtpWebRequest]::Create($FTPLink)
	$FTPRequest.set_Timeout(10000) #10 second timeout
	$FTPRequest.Credentials = New-Object System.Net.NetworkCredential($FTPUser, $FTPPassword)
    $FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::GetFileSize
    $FTPRequest.UseBinary = $true
	$FTPRequest.KeepAlive = $false
    $FTPResponse = $FTPRequest.GetResponse()
    $FileLength = [System.Math]::Floor($FTPResponse.ContentLength/1024/1024)
    $FTPResponse.Close()
    return $FileLength                     
}

#Download file from FTP
function DownloadFileFromFTP([String] $FTPUser, [String] $FTPPassword, [String] $FTPLink, [String] $DSTFloder)
{
	if( [String]::IsNullOrEmpty($FTPLink))
    {
        Write-Host "Download Link is empty. Skip the download job."
		return ""
    }

    if(!(Test-Path $DSTFloder))
    {
        try
        {
            $null = New-Item -ItemType Directory $DSTFloder -ErrorAction Stop
        }
        catch
        {
            $NewDSTFloder = "C:"
            Write-Host "Local path '$DSTFloder' cannot be created/reached. File will be save in $NewDSTFloder." -ForegroundColor "Yellow"
            $DSTFloder = $NewDSTFloder
        }
    }	
    $FileName = $FTPLink.Substring($FTPLink.LastIndexOf("/") + 1)
	$DownloadedFile = "$DSTFloder\$FileName"
	try
	{
		Write-Host "Start downloading $FTPLink to $DownloadedFile."
		$FTPRequest = [System.Net.FtpWebRequest]::Create($FTPLink)
		$FTPRequest.set_Timeout(10000) #10 second timeout
		$FTPRequest.Credentials = New-Object System.Net.NetworkCredential($FTPUser, $FTPPassword)
        $FTPRequest.Method = [System.Net.WebRequestMethods+Ftp]::DownloadFile
        $FTPRequest.UseBinary = $true
		$FTPRequest.KeepAlive = $false
        $FTPResponse = $FTPRequest.GetResponse()
		$ResponseStream = $FTPResponse.GetResponseStream()
        #Get the file size
        $FileLength = Get-FTPFilesize $FTPUser $FTPPassword $FTPLink
		$TargetStream = New-Object IO.FileStream ($DownloadedFile, [IO.FileMode]::Create)
		[byte[]]$ReadBuffer = New-Object byte[] 10KB

        $Count = $ResponseStream.Read($ReadBuffer,0,$ReadBuffer.length)
        $DownloadedBytes = $Count

        while ($Count -ne 0)
        {
            $TargetStream.Write($ReadBuffer, 0, $Count)

            $Count = $ResponseStream.Read($ReadBuffer,0,$ReadBuffer.length)

            $DownloadedBytes += $Count

            Write-Progress -activity "Downloading file $FileName to $DSTFloder" -status "Downloaded ($([System.Math]::Floor($DownloadedBytes/1024/1024))M of $($FileLength)M): " -PercentComplete ((([System.Math]::Floor($DownloadedBytes/1024/1024)) / $FileLength)  * 100)
        }

        Write-Progress -activity "Finished downloading file $FileName to $DSTFloder"
        $targetStream.Flush()
        $targetStream.Close()
        $targetStream.Dispose()
        $responseStream.Dispose()
        return $DownloadedFile
	}
	catch
	{
		Write-Host "Error occurred while downloading from $FTPLink." -ForegroundColor "Red"
		Write-Host "$_" -ForegroundColor "Red"
		Write-Host "Please check settings, such as ftp account, password, download link, local folder, etc, and download manually." -ForegroundColor "Red"
        return ""
	}
}

$FTPUser = "dbm"
$FTPLink = "ftp://software.databasemart.net/Software_System/MSSQL/$SQLFile"
$DSTFloder = "C:/"

Write-Host "Downloading $SQLFile from FTP server..." -ForegroundColor Green
#Call the function to download the file
$SQLFile = DownloadFileFromFTP  $FTPUser $FTPPassword $FTPLink $DSTFloder

if($SQLVersion -like "SQL_2016*")
{
    $SPFile = "sqlserver2016sp2-kb4052908-x64-enu.exe"
    Write-Host "Downloading $SPFile from FTP server..." -ForegroundColor Green
    $URL = "ftp://software.databasemart.net/Software_System/MSSQL/$SPFile"
    $SPFile = DownloadFileFromFTP  $FTPUser $FTPPassword $URL $DSTFloder
}
elseif($SQLVersion -like "SQL_2014*")
{
    #install .net 3.5
    DISM /Online /Enable-Feature /FeatureName:NetFx3 /All | Out-Null
    $SPFile = "sqlserver2014sp3-kb4022619-x64-enu.exe"
    Write-Host "Downloading $SPFile from FTP server..." -ForegroundColor Green
    $URL = "ftp://software.databasemart.net/Software_System/MSSQL/$SPFile"
    $SPFile = DownloadFileFromFTP  $FTPUser $FTPPassword $URL $DSTFloder
}
elseif($SQLVersion -like "SQL_2012*")
{
    $SPFile = "sqlserver2012sp4-kb4018073-x64-enu.exe"
    Write-Host "Downloading $SPFile from FTP server..." -ForegroundColor Green
    $URL = "ftp://software.databasemart.net/Software_System/MSSQL/$SPFile"
    $SPFile = DownloadFileFromFTP  $FTPUser $FTPPassword $URL $DSTFloder
}
elseif($SQLVersion -like "SQL_2008*")
{
    $SPFile = "sqlserver2008sp4-kb2979596-x64-enu.exe"
    Write-Host "Downloading $SPFile from FTP server..." -ForegroundColor Green
    $URL = "ftp://software.databasemart.net/Software_System/MSSQL/$SPFile"
    $SPFile = DownloadFileFromFTP  $FTPUser $FTPPassword $URL $DSTFloder
}

Write-Host "Expanding $SQLFile..." -ForegroundColor Green
#Expand the Zip file
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($SQLFile, "C:/")

Write-Host "Installing $SQLVersion..." -ForegroundColor Green
#Run install.bat
Set-Location C:/$SQLVersion
$installProsess = Start-Process ./install.bat -Wait -Passthru
$installProsess.WaitForExit()
if($installProsess.ExitCode -eq 0){
    Write-Host "Install complete!" -ForegroundColor Green
}

#load the SMO dll, and connect to the SQL Server
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
$sql = new-object ('Microsoft.SqlServer.Management.Smo.Server') 'localhost'
#Change to Mixed Mode
$sql.Settings.LoginMode = [Microsoft.SqlServer.Management.SMO.ServerLoginMode] 'Mixed'
#Change maximum memory
$sql.Configuration.MaxServerMemory.ConfigValue = 4096
# Make the changes
$sql.Alter()
$saLogin = $sql.Logins.Item('sa')
#Enable sa login
$saLogin.Enable()
$saLogin.PasswordPolicyEnforced  = $False 
$saLogin.Alter()
#Change sa password
$saLogin.ChangePassword($saPwd)
$saLogin.Alter()

# Restart SQL Server to apply changes
Restart-Service -Name MSSQLSERVER

#Enabling SQL Server Ports
New-NetFirewallRule -DisplayName “SQL Server” -Direction Inbound –Protocol TCP –LocalPort 1433 -Action allow | Out-Null
New-NetFirewallRule -DisplayName “WhiteList” -Direction Inbound -Action allow -RemoteAddress 74.124.24.0/24,172.106.164.0/25 | Out-Null

Set-Location C:/
if($SPFile)
{
    Write-Host "Installing SQL Server update package..." -ForegroundColor Green
    $installUpdateProsess = Start-Process $SPFile -ArgumentList "/allinstances","/qs","/IAcceptSQLServerLicenseTerms" -Wait -Passthru
    $installUpdateProsess.WaitForExit()
    if($?){
        Write-Host "Install complete!" -ForegroundColor Green
        Remove-Item $SPFile -Recurse -Force
    }
}
#Delete files
Remove-Item C:/$SQLVersion,$SQLFile -Recurse -Force

#Start SQL Server service
Start-Service -Name MSSQLSERVER

#Windows OS Version
$OSVersion = Get-WmiObject -Class Win32_OperatingSystem | Select-Object -ExpandProperty Caption
if ($OSVersion.Contains("2016") -or $OSVersion.Contains("2019")) {
    #Remove from recent files
    (New-Object -ComObject shell.application).Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}").Items() | ?{$_.Path -like "*InstallSQL*"} | % {$_.InvokeVerb("remove")}
}

#Delete the .ps1 file and close the window
#Remove-Item -Force (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Path
#Get-Process -Id $PID | Stop-Process

