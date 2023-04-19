<#


                                           .       .x+=:.                                          .         s               
   ..                                     @88>    z`    ^%                               oec :    @88>      :8               
  888>                 ..    .     :      %8P        .   <k        u.      u.    u.     @88888    %8P      .88         ..    
  "8P         u      .888: x888  x888.     .       .@8Ned8"  ...ue888b   x@88k u@88c.   8"*88%     .      :888ooo    .@88i   
   .       us888u.  ~`8888~'888X`?888f`  .@88u   .@^%8888"   888R Y888r ^"8888""8888"   8b.      .@88u  -*8888888   ""%888>  
 u888u. .@88 "8888"   X888  888X '888>  ''888E` x88:  `)8b.  888R I888>   8888  888R   u888888> ''888E`   8888        '88%   
`'888E  9888  9888    X888  888X '888>    888E  8888N=*8888  888R I888>   8888  888R    8888R     888E    8888      ..dILr~` 
  888E  9888  9888    X888  888X '888>    888E   %8"    R88  888R I888>   8888  888R    8888P     888E    8888     '".-%88b  
  888E  9888  9888    X888  888X '888>    888E    @8Wou 9%  u8888cJ888    8888  888R    *888>     888E   .8888Lu=   @  '888k 
  888E  9888  9888   "*88%""*88" '888!`   888&  .888888P`    "*888*P"    "*88*" 8888"   4888      888&   ^%888*    8F   8888 
  888E  "888*""888"    `~    "    `"`     R888" `   ^"F        'Y"         ""   'Y"     '888      R888"    'Y"    '8    8888 
  888E   ^Y"   ^Y'                         ""                                            88R       ""             '8    888F 
  888E                                                                                   88>                       %k  <88F  
  888P                                                                                   48                         "+:*%`   
.J88" "                                                                                  '8                                  


.SYNOPSIS
    PowerShell script to automate the process of updating the WebDriver and ChromeDriver used for Selenium.

.DESCRIPTION
    PowerShell script to help automate the process of updating the WebDriver library and ChromeDriver executable used for browser automation testing with Selenium. 
    Uses Google Chrome's version to match the correct ChromeDriver.exe. Also downloads the latest version of WebDriver. 
    Makes backups and logs changes.

.PARAMETER  None

.EXAMPLE
    C:\PS> .\ChromeDriverUpdater.ps1

.NOTES
    Author: JamisonFitz
    Date: 4/14/2023
    Version: 1.0
    License: Apache-2.0 license

.LINK
    GitHub repository: https://github.com/your-username/ChromeDriverUpdater

#>

# License information
<#
Copyright (c) [2023] [JamisonFitz]
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

cls

# Change to your preferred path if needed; the path below is the user's home directory
$destinationPath = "$env:USERPROFILE"

$webDriverPath = Join-Path $destinationPath 'WebDriver.dll'
$chromeDriverPath = Join-Path $destinationPath 'ChromeDriver.exe'
$webDriverPathbackup = Join-Path $destinationPath 'Backup_WebDriver.dll'
$chromeDriverPathbackup = Join-Path $destinationPath 'Backup_ChromeDriver.exe'

if (Test-Path $webDriverPath) {
    $oldwebDriverVersion = (Get-Item $webDriverPath).VersionInfo.FileVersion
    if (Test-Path $webDriverPathbackup) {
        Remove-Item $webDriverPathbackup -Force
    }
    else {
        Rename-Item -Path $webDriverPath -NewName 'Backup_WebDriver.dll' -Force
    }
}

if (Test-Path $chromeDriverPath) {
    $oldchromeDriverVersion = (& $chromeDriverPath --version).Split(' ')[1] 
    if (Test-Path $chromeDriverPathbackup) {
        Remove-Item $chromeDriverPathbackup -Force
    }
    else {
        Rename-Item -Path $chromeDriverPath -NewName 'Backup_ChromeDriver.exe' -Force
    }
}

$chromeLocations = @(
    'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome',
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "$env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)

$chromeVersion = $null

foreach ($location in $chromeLocations) {
    $version = (Get-Item $location -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
    if ($version) {
        $chromeVersion = $version
        break
    }
}

if ($chromeVersion) {
    Write-Host "Google Chrome version: $chromeVersion"
}
else {
    Write-Host "Google Chrome not found"
}

# Construct the URL to get the latest release version of ChromeDriver
$chromeDriverUrl = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_$($chromeVersion.Split('.')[0])"
$chromeDriverVersion = Invoke-WebRequest -Uri $chromeDriverUrl | Select-Object -ExpandProperty Content

$chromeDriverUrl = "https://chromedriver.storage.googleapis.com/$chromeDriverVersion/chromedriver_win32.zip"

# Download and extract ChromeDriver
Invoke-WebRequest -Uri $chromeDriverUrl -OutFile (Join-Path $destinationPath 'chromedriver.zip')

if ((Get-Item (Join-Path $destinationPath 'chromedriver.zip')).Extension -ne '.zip') {
    Write-Error "Failed to download ChromeDriver package"
    exit
}

Expand-Archive -Path (Join-Path $destinationPath 'chromedriver.zip') -DestinationPath $destinationPath -Force
Remove-Item (Join-Path $destinationPath 'chromedriver.zip')

$chromeDriverPath = Join-Path $destinationPath 'chromedriver.exe'

Write-Host "$chromeDriverPath Was: $oldchromeDriverVersion Now: $chromeDriverVersion"

$rng = Get-Random -Minimum 0 -Maximum 999
# Update WebDriver
$webDriverUrl = "https://www.nuget.org/api/v2/package/Selenium.WebDriver"
$zipPath = Join-Path ([System.IO.Path]::GetTempPath()) "$rng-selenium.zip"
$extractPath = Join-Path ([System.IO.Path]::GetTempPath()) "$rng-selenium"

try {
    Invoke-WebRequest -Uri $webDriverUrl -OutFile $zipPath
    Write-Host 'Updating WebDriver to the latest version...'
    Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force

    $dllPath = Join-Path $extractPath 'lib\net45\WebDriver.dll'
    # Swap net45 for net47, net48, net5.0, net6.0, netstandard2.0, or netstandard2.1
    Copy-Item -Path $dllPath -Destination $destinationPath -Force

}
catch {
    Write-Host "WebDriver update failed with error: $_"
    exit
}

Write-Host "$webDriverPath Was: $oldwebDriverVersion Now: $webDriverVersion_new"

#Log changes
$logFilePath = Join-Path $destinationPath 'WebDriverUpdater.log'

if (-not (Test-Path $logFilePath)) {
    New-Item -ItemType File -Path $logFilePath -Force | Out-Null
}

$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$logMessage = "{0} - WebDriver updated to the latest version. ChromeDriver upgraded from {1} to {2}" -f $timestamp, $oldchromeDriverVersion, (& $chromeDriverPath --version).Split(' ')[1]
$logMessage += [Environment]::NewLine + "{0} - WebDriver updated to the latest version. WebDriver.dll upgraded from {1} to {2}" -f $timestamp, $oldwebDriverVersion, $webDriverVersion_new

Add-Content -Path $logFilePath -Value $logMessage
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$logMessage = "{0} - WebDriver updated to the latest version. ChromeDriver upgraded from {1} to {2}" -f $timestamp, $oldchromeDriverVersion, (& $chromeDriverPath --version).Split(' ')[1]
$logMessage += [Environment]::NewLine + "{0} - WebDriver updated to the latest version. WebDriver.dll upgraded from {1} to {2}" -f $timestamp, $oldwebDriverVersion, $webDriverVersion_new

Add-Content -Path $logFilePath -Value $logMessage