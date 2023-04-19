# ChromeDriver_WebDriver_Refresher

This PowerShell script helps automate the process of updating the WebDriver library and ChromeDriver executable used for browser automation testing with Selenium. It uses Google Chrome's version to match the correct ChromeDriver.exe and downloads the latest version of WebDriver. The script also makes backups of the existing versions and logs the update process.

Customization
By default, the script installs the updated executables in the user's home directory. However, you can change the $destinationPath variable to specify a different installation directory.

Additionally, you can customize which version of WebDriver you want to install by changing the net45 directory in the $webDriverUrl variable to a different version, such as net48 or netstandard2.1.

Requirements
This script requires PowerShell to run and relies on the following external dependencies:

Google Chrome web browser (could be modified for others)
Internet connection to download the latest versions of ChromeDriver and Selenium WebDriver

Usage
To use this script, simply download the WebDriverUpdater.ps1 file and run it in a PowerShell terminal. You may need to set your PowerShell execution policy to allow script execution by running Set-ExecutionPolicy RemoteSigned or Set-ExecutionPolicy Unrestricted. After running the script, the updated versions of WebDriver and ChromeDriver will be installed in the specified directory, and a log file will be created to document the update process.

License
This script is licensed under the MIT License. Feel free to use and modify it as needed.
