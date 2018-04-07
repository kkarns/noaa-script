#######################################################################################################################
## 
## name:    
##      noaa-script.ps1
##
##      daily download of *.DAT data files on the noaa ftp server.
##
## syntax:
##      .\noaa-script.ps1
##
## dependencies:
##  Depends on a specific unique file structure on the vendor's ftp server.  Code is solving this unique server problem only.
##  C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe.config    -- for dotnet 4 library support on older PS environments
##  credentials.txt         -- from:  read-host -assecurestring | convertfrom-securestring | out-file C:\path-to-app\credentials.txt
##  $MyDir\Settings.xml     -- an XML settings file
##
## updated:
##      2017-04-10 original version 
##      2018-04-07 pulled onto new client/servergit --version
##
## todo:
##      fix the case where, in January, we are looking back into a prior year
##

## Functions ##########################################################################################################

##
## LogWrite - write messages to log file 
##

Function LogWrite
{
   Param ([string]$logstring)
   Add-content $Logfile -value $logstring 
}

##
## Get-FtpDir - Connect and Download from http://stackoverflow.com/questions/19059394/powershell-connect-to-ftp-server-and-get-files
##

Function Get-FtpDir ($url,$credentials) {
    $request = [Net.WebRequest]::Create($url)
    $request.Method = [System.Net.WebRequestMethods+FTP]::ListDirectory
    if ($credentials) { $request.Credentials = $credentials }
    $response = $request.GetResponse()
    $reader = New-Object IO.StreamReader $response.GetResponseStream() 
    while(-not $reader.EndOfStream) {
        $reader.ReadLine()
    }
    #$reader.ReadToEnd()
    $reader.Close()
    $response.Close()
}

## Main Code ##########################################################################################################

try {

##
## Import settings from XML config file
##    variables format is: $ConfigFile.noaaScript.x
##    note: XML config file needs to be well formed, verify it in IE first.
##
$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
[xml]$ConfigFile = Get-Content "$myDir\Settings.xml"

## setup the logfile
$LogDir = $myDir + "\logs"
if(-not ([IO.Directory]::Exists($LogDir))) {New-Item -ItemType directory -Path $LogDir}
$Logfile = ($LogDir + "\noaa-script-" + $(get-date -f yyyy-MM-dd-HHmmss) + ".log")
echo "results are logged to:  "$Logfile 
LogWrite ("Started at:  " + $(get-date -f yyyy-MM-dd-HHmmss))
$date1 = Get-Date

## Load the dotnet assembly for sftp from NUGET
##[Void][Reflection.Assembly]::LoadFrom("$MyDir\Renci.SshNet.dll")        

##
## Get the servername and username from the settings.xml file 
##
$serverName = $ConfigFile.noaaScript.noaaServerName
$userName = $ConfigFile.noaaScript.noaaUserName
LogWrite ("servername   :  " + $serverName)
LogWrite ("username     :  " + $userName)

##
## Get the directory information from the settings.xml file 
##

## local
##
##$datFolder = "D:\data\noaa\"
##
$datFolder = $ConfigFile.noaaScript.datFolder
LogWrite ("datFolder  :  " + $datFolder)
if(-not ([IO.Directory]::Exists($datFolder))) 
    {
    echo ("Error. Halted. Couldn't find dat source directory on this server.")
    LogWrite ("Error. Halted. Couldn't find dat source directory on this server.")
    throw ("Error. Halted. Couldn't find dat source directory on this server.")
    }

## remote
##
##noaaRemoteDirectory      = "/data/radiation/surfrad/Boulder_CO"
##noaaRemoteDirectoryFinal = "/data/radiation/surfrad/Boulder_CO/2017"
##
$noaaRemoteDirectory = $ConfigFile.noaaScript.noaaRemoteDirectory
LogWrite ("noaaRemoteDirectory           :  " + $noaaRemoteDirectory)


##
## Get the password from the encrypted credentials file 
## 
## https://blogs.technet.microsoft.com/robcost/2008/05/01/powershell-tip-storing-and-using-password-credentials/
## note the pre-requisite (as explained in the blog)
##     credentials.txt   
##         from:  
##             read-host -assecurestring | convertfrom-securestring | out-file D:\scripts\noaa\credentials.txt
##
$credentialsFile = $MyDir+ "\" + $ConfigFile.noaaScript.noaaCredentialsFile
if(![System.IO.File]::Exists($credentialsFile))
    {
    echo ("Error. Halted. Missing encrypted credentials file.")
    LogWrite ("Error. Halted. Missing encrypted credentials file.")
    throw ("Error. Halted. Missing encrypted credentials file.")
    }
$password = get-content $credentialsFile | convertto-securestring
$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $userName,$password
LogWrite ("password          :  " + $password)
LogWrite ("credentials       :  " + $credentials)
LogWrite ("decrypted username:  " + $credentials.GetNetworkCredential().UserName)
LogWrite ("decrypted password:  " + $credentials.GetNetworkCredential().password)
$passwordInFile = $credentials.GetNetworkCredential().password


##
## look back a week for any new files at noaa - source: http://stackoverflow.com/questions/19059394/powershell-connect-to-ftp-server-and-get-files
##

$ftp = $ConfigFile.noaaScript.noaaFtp                   ## "ftp://aftp.cmdl.noaa.gov"
$user = $ConfigFile.noaaScript.noaaUserName             ## 'anonymous'
$pass = $passwordInFile                                 ## see above
$folder = $ConfigFile.noaaScript.noaaRemoteDirectory    ## need to add the year to the path 'data/radiation/surfrad/Boulder_CO/2017'  
$target = $ConfigFile.noaaScript.datFolder              ## "D:\data\noaa\"

#set credentials
$credentials = new-object System.Net.NetworkCredential($user, $pass)

#set folder path
$folder = $folder + "/" + $(get-date -f yyyy).ToString()    ## add the year to the path 'data/radiation/surfrad/Boulder_CO/2017' 
$folderPath= $ftp + "/" + $folder + "/"

$files = Get-FTPDir -url $folderPath -credentials $credentials
#$files              #uncomment this line to show the list of all the filenames in the remote ftp server

$webclient = New-Object System.Net.WebClient 
$webclient.Credentials = New-Object System.Net.NetworkCredential($user,$pass) 

ForEach ($i in 7..1) {
    $a = Get-Date
    $d = $a.AddDays(-1 * $i)
    "looking for file from:       " + $d
    $file = "tbl" + ($d.Year % 100).ToString() + $d.DayOfYear.ToString().PadLeft(3,'0') + ".dat"    
    $source=$folderPath + $file  
    $destination = $target + $file 
    "ready to download.  source:  " + $source + "  dest:  " + $destination
    LogWrite ("ready to download.  source:  " + $source + "  dest:  " + $destination)       
    try {
        $webclient.DownloadFile($source, $target+$file)
    } 
    Catch {
        ##
        ## log any error
        ##    
        "    skipping, failed to find file.  source:  " + $source 
        LogWrite ("    skipping, failed to find file.  source:  " + $source )
        LogWrite $Error[0]
    }
    echo "---------------------"
}


throw ("Halted.  This is the end.  Who knew.")


}
Catch {
    ##
    ## log any error
    ##    
    LogWrite $Error[0]
}
Finally {

    ##
    ## go back to the software directory where we started
    ##
    set-location $myDir

    LogWrite ("finished at:  " + $(get-date -f yyyy-MM-dd-HHmmss))
}