#Created by Trevor Faust
#Modified: 5-9-2018

Write-Host "Start Script"
Write-Host ((Get-date).ToShortTimeString())

Write-Host "Package Name is: 7-Zip"
Write-Host "Package Version is 1805"

#Script uses commands that do NOT work on Powershell Versions older than 4.
#Set this to Version 5 to ensure our environment has the latest Powershell Version.
Write-Host "Checking PowerShell Version being 5 or later"

$PowerShellVersion = $PSVersionTable.PSVersion.Major

Write-Host "PowerShell Version: $PowerShellVersion"
if ($PowerShellVersion -lt 5)
{
	Write-Host "PowerShell Version is less than 5, Exiting Script"
	
	Write-Host "End Script at"
	Write-Host ((Get-Date).ToShortDateString())
	exit
}

# Function for the Installation of 7-Zip
function Install-7Zip
{
	Write-Host "`nBegin Installation of 7-Zip Version 1805"
	
	if ($Script:OSArch -eq $true)
	{
		Write-Host "Installing for 64-Bit Operating System"
		Write-Host "Running Command: $env:SystemRoot\Sysnative\cmd.exe /C msiexec.exe /i 7z1805-x64.msi /norestart /qn /L*v $env:TEMP\7-zip-1805.log"
		try
		{
            #Runs cmd from sysnative for 64-bit system then spawns msiexec with the install argument for 7-zip
			$InstallProc64 = Start-Process -FilePath "$env:SystemRoot\Sysnative\cmd.exe" -ArgumentList " /C msiexec.exe /i 7z1805-x64.msi /norestart /qn /L*v $env:TEMP\7-zip-1805.log" -NoNewWindow -PassThru -Wait -ErrorAction Stop
			Write-Host "Successfully Installed 7-zip!"
			Write-Host "Exit Code: $($InstallProc64.ExitCode)"
		}
		catch [exception]
		{
			Write-Host "Failed to Install 7-Zip`n`tERROR: $_"
			Write-Host "`n`nExit Code: $($InstallProc.ExitCode)"
			Write-Host "Please reivew Logs located ($env:TEMP\7-zip-1805.log)...`n`nExiting Script."
			Write-Host "End Script at"
			Write-Host ((Get-Date).ToShortDateString())
			exit
		}
	}
	else
	{
		Write-Host "Installing for 32-Bit Operating System"
		Write-Host "Running Command: msiexec.exe /i 7z1805.msi /norestart /qn /L*v $env:TEMP\7-zip-1805.log"
		
		try
		{
            #Runs cmd from System32 for 32-bit system then spawns msiexec with the install argument for 7-zip
			$InstallProc = Start-Process -FilePath "$env:SystemRoot\System32\cmd.exe" -ArgumentList " /C msiexec.exe /i 7z1805.msi /norestart /qn /L*v $env:TEMP\7-zip-1805.log" -NoNewWindow -PassThru -Wait -ErrorAction Stop
			Write-Host "Successfully Installed 7-zip!"
			Write-Host "Exit Code: $($InstallProc.ExitCode)"
		}
		catch [exception]
		{
			Write-Host "Failed to Install 7-Zip`n`tERROR: $_"
			Write-Host "`n`nExit Code: $($InstallProc.ExitCode)"
			Write-Host "Please reivew Logs located ($env:TEMP\7-zip-1805.log)...`n`nExiting Script."
			Write-Host "End Script at"
			Write-Host ((Get-Date).ToShortDateString())
			exit
		}
		
	}
	
}

#Gets Architecture of a system being 64-bit (Returns True for 64-bit and False for 32-Bit)
$script:OSArch = [Environment]::Is64BitOperatingSystem

if ($script:OSArch -eq $true)
{
    #Checks 64-bit Registry Uninstall Strings
	$key64 = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
	$subKey64 = $key64.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
	Write-Host "System Architecture is 64-Bit"
	Write-Host "`nWill uninstall everything pertaining to 7-Zip from 64 and 32-bit Uninstall registries."
}
else
{
	Write-Host "System Architecture is 32-Bit"
	Write-Host "`nWill uninstall everything pertaining to 7-Zip from the Uninstall Registry"
}

$Found = 0
#Checks 32-bit Registry (For 64-bit systems) or just the native registry (for 32-bit systems) for Uninstall Strings
$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry32)
$subKey = $key.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")

#64-bit Registry Parsing
if ($subKey64 -ne $null)
{
	foreach ($Key64 in $subKey64.GetSubKeyNames())
	{
        #Checks the Display name of all Uninstall keys for 7-zip
		if ($subKey64.OpenSubKey($Key64).GetValue("DisplayName") -like "*7-zip*")
		{
			$Found = 1
			#Found 7-Zip, begin uninstall procedure
			Write-Host "Found $($subKey64.OpenSubKey($Key64).GetValue("DisplayName")) in 64bit Registry`n"
			Write-Host ($subKey64.OpenSubKey($Key64).GetValue("UninstallString"))
			Write-Host ($subKey64.OpenSubKey($Key64).GetValue("DisplayVersion"))
			
			#Uninstall Application
			Write-Host "`nUninstalling 7-Zip"
			
			try
			{
                #Matches for a "{" because that means it's an MSI and needs to run msiexec.exe /x
				if (($subKey64.OpenSubKey($Key64).GetValue("UninstallString")) -match "{")
				{
                    #Extracts the GUID from the uninstall string.
					$GUID = (($subKey64.OpenSubKey($Key64).GetValue("UninstallString")) -replace "^msiexec.exe (/x|/i)").Trim()
                    #Executes msiexec.exe /x with the extracted GUID from above
					Write-Host "Running Command: $env:SystemRoot\Sysnative\cmd.exe /C msiexec.exe /X $GUID /qn /norestart"
					$UninstallProc64 = Start-Process -FilePath "$env:SystemRoot\Sysnative\cmd.exe" -ArgumentList " /C msiexec.exe /X $GUID /qn /norestart" -NoNewWindow -PassThru -Wait -ErrorAction Stop
				}
				else
				{
                    Write-Host "Running Command: $env:SystemRoot\Sysnative\cmd.exe /C ""$($subKey64.OpenSubKey($Key64).GetValue("UninstallString") -replace '"')"" /S"
                    #Uninstall String is NOT an msi command. Run the uninstall String with /S as the argument for silently uninstalling the product.
					$UninstallProc64 = Start-Process -FilePath "$env:SystemRoot\Sysnative\cmd.exe" -ArgumentList " /C ""$($subKey64.OpenSubKey($Key64).GetValue("UninstallString") -replace '"')"" /S" -NoNewWindow -PassThru -Wait -ErrorAction Stop
				}
				Write-Host "7-Zip Uninstalled Successfully!"
				Write-Host "Exit Code: $($UninstallProc64.ExitCode)"
			}
			catch [exception]
			{
				Write-Host "Failed to Uninstall 7-Zip`n`tERROR: $_"

				Write-Host "End Script at"
				Write-Host ((Get-Date).ToShortDateString())
				exit
			}
		}
	}
}

#32-bit Registry Parsing
foreach ($Key in $subKey.GetSubKeyNames())
{
	if ($subKey.OpenSubKey($Key).GetValue("DisplayName") -like "*7-zip*")
	{
		$Found = 1
		
		Write-Host "Found $($subKey.OpenSubKey($Key).GetValue("DisplayName")) in 32bit Registry`n"
		Write-Host ($subKey.OpenSubKey($Key).GetValue("UninstallString"))
		Write-Host ($subKey.OpenSubKey($Key).GetValue("DisplayVersion"))
		
		#Uninstall Application
		Write-Host "`nUninstalling 7-Zip"
		
        try
		{
            #Matches for a "{" because that means it's an MSI and needs to run msiexec.exe /x
			if (($subKey.OpenSubKey($Key).GetValue("UninstallString")) -match "{")
			{
                #Extracts the GUID from the uninstall string.
				$GUID = (($subKey.OpenSubKey($Key).GetValue("UninstallString")) -replace "^msiexec.exe (/x|/i)").Trim()
                #Executes msiexec.exe /x with the extracted GUID from above
				Write-Host "Running Command: $env:SystemRoot\System32\cmd.exe /C msiexec.exe /X $GUID /qn /norestart"
				$UninstallProc = Start-Process -FilePath "$env:SystemRoot\System32\cmd.exe" -ArgumentList " /C msiexec.exe /X $GUID /qn /norestart" -NoNewWindow -PassThru -Wait -ErrorAction Stop
			}
			else
			{
                Write-Host "Running Command: $env:SystemRoot\System32\cmd.exe /C ""$($subKey.OpenSubKey($Key).GetValue("UninstallString") -replace '"')"" /S"
                #Uninstall String is NOT an msi command. Run the uninstall String with /S as the argument for silently uninstalling the product.
				$UninstallProc = Start-Process -FilePath "$env:SystemRoot\System32\cmd.exe" -ArgumentList " /C ""$($subKey.OpenSubKey($Key).GetValue("UninstallString") -replace '"')"" /S" -NoNewWindow -PassThru -Wait -ErrorAction Stop
			}
			Write-Host "7-Zip Uninstalled Successfully!"
			Write-Host "Exit Code: $($UninstallProc.ExitCode)"
		}
		catch [exception]
		{
			Write-Host "Failed to Uninstall 7-Zip`n`tERROR: $_"

			Write-Host "End Script at"
			Write-Host ((Get-Date).ToShortDateString())
			exit
		}
	}
}

if ($Found -eq 0)
{
	Write-Host "7-Zip Is not Installed!`n`n No Uninstall Registry value for 7-Zip... Skipping to Installer!"
}

Install-7Zip

Write-Host "End Script at"
Write-Host ((Get-Date).ToShortDateString())
