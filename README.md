# Useful PowerShell Commands Collection

## Software and System Management

### Package Manager (Chocolatey)
Install Chocolatey package manager:
```powershell
@"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"
```

Search for available packages:
```powershell
choco search <package-name>
```

Install a package:
```powershell
choco install <package-name> -y
```

List installed packages:
```powershell
choco list --local-only
```

### Deploy UltraVNC
Silent installation and configuration of UltraVNC:
```powershell
UltraVNC_1_2_40_X64_Setup.exe /silent /loadinf="C:\temp\config.inf"
xcopy /Y ultravnc.ini "C:\Program Files\uvnc bvba\UltraVNC\"
"C:\Program Files\uvnc bvba\UltraVNC\winvnc.exe" /silent
"C:\Program Files\uvnc bvba\UltraVNC\unins000.exe" /silent
net stop uvnc_service
net start uvnc_service
```

### List Installed Programs
Get a list of all installed programs:
```powershell
Get-WmiObject -Query "SELECT * FROM Win32_Product" | Select-Object -Property Name
```

Get detailed information about installed programs (including version, publisher, and install date):
```powershell
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
```

### Uninstall Software
Uninstall a specific program using command line:
```powershell
wmic product where "name='NomeProgramma'" call uninstall
```

Uninstall SAP silently:
```powershell
"C:\Program Files (x86)\SAP\SapSetup\Setup\NwSapSetup.exe" /all /Silent /uninstall
```

### Group Policy Reports
Generate a text report of Group Policy settings for a specific user:
```powershell
gpresult /user DOMINIO\Mario /r
```

Generate an HTML report of Group Policy settings:
```powershell
gpresult /user dominio\user /h report.html
```

### Font Installation
Install all TTF fonts from a specified folder:
```powershell
$fontFolder = "\\ritsrvvshr01\rubdata\Utilities\Preparazione PC\Open_Sans\static"  # Percorso condiviso
$tempFolder = "C:\Temp"

# Inizia la registrazione
Start-Transcript -Path "C:\Temp\FontInstallLog.txt" -Append

# Scarica i font dal percorso condiviso al temp
Copy-Item -Path $fontFolder\*.ttf -Destination $tempFolder -Force

$fontFiles = Get-ChildItem -Path $tempFolder -Filter *.ttf

foreach ($font in $fontFiles) {
    $fontPath = $font.FullName
    $fontName = $font.Name
    Write-Host "Verificando: $fontName"

    $destinationPath = Join-Path -Path "$env:SystemRoot\Fonts" -ChildPath $fontName

    if (Test-Path -Path $destinationPath) {
        Write-Host "$fontName è già installato."
        Write-EventLog -LogName Application -Source "FontInstallScript" -EventID 1001 -EntryType Information -Message "$fontName è già installato."
    } else {
        Write-Host "Installando: $fontName"
        Copy-Item -Path $fontPath -Destination $destinationPath -Force

        $regPath = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
        Set-ItemProperty -Path $regPath -Name $fontName -Value $fontName
        
        Write-Host "Installazione completata per: $fontName"
        Write-EventLog -LogName Application -Source "FontInstallScript" -EventID 1002 -EntryType Information -Message "Installazione completata per: $fontName"
    }
}

Write-Host "Tutti i font sono stati verificati e installati!"
Write-EventLog -LogName Application -Source "FontInstallScript" -EventID 1003 -EntryType Information -Message "Tutti i font sono stati verificati e installati!"

# Termina la registrazione
Stop-Transcript

```

### Power Management
Disable hibernation on the system:
```powershell
powercfg /h off
```

## Hardware Information

### USB Devices
List all USB devices with detailed information:
```powershell
wmic path CIM_LogicalDevice where "Description like 'USB%'" get /value
```

### Monitor Information
Get monitor details including manufacturer, model, and serial number:
```powershell
Get-WmiObject WmiMonitorID -Namespace root\wmi | Select-Object `
@{Name="Manufacturer";Expression={[System.Text.Encoding]::ASCII.GetString($_.ManufacturerName)}}, `
@{Name="Model";Expression={[System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName)}}, `
@{Name="SerialNumber";Expression={[System.Text.Encoding]::ASCII.GetString($_.SerialNumberID)}}
```

## System Administration

### Windows Update
Scan for Windows updates:
```powershell
UsoClient ScanInstallWait
```

### Active Directory
Get detailed information about a specific computer:
```powershell
Get-ADComputer -Identity "RUB0138" -Properties *
```

Export all Active Directory users with their OUs:
```powershell
$utenti = Get-ADUser -Filter * -Property DisplayName, Name, Surname, SamAccountName, EmailAddress, DistinguishedName

# Seleziona le proprietà desiderate e isola le OU
$utentiSelezionati = $utenti | Select-Object DisplayName, Name, Surname, SamAccountName, EmailAddress,
    @{Name='OU'; Expression={ 
        # Estrai il DistinguishedName
        $dn = $_.DistinguishedName
        
        # Split il DN in base alla virgola e filtra solo le OU
        $ouList = $dn -split ',' | Where-Object { $_ -like 'OU=*' }
        
        # Restituisci le OU come una stringa separata da un punto e virgola
        return ($ouList -join '; ')
    }}

# Esporta in un file CSV
$utentiSelezionati | Export-Csv -Path "C:\temp\UtentiAD.csv" -NoTypeInformation -Encoding UTF8
```

### Microsoft Office Information
Get Office Click-to-Run configuration details:
```powershell
Get-ItemProperty -Path "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration" | Select-Object -Property ProductReleaseIds, ProductVersion
```

### System Maintenance
Schedule a system restart with a 2-hour delay and warning message:
```powershell
shutdown /r /t 7200 /c "Il sistema verrà riavviato tra 2 ore per manutenzione. Si prega di salvare il lavoro in corso."
```

### User Profile Management
List all user profiles:
```powershell
Get-WmiObject -Class Win32_UserProfile | Select-Object LocalPath
```

Delete a specific user profile:
```powershell
$profile = Get-WmiObject -Class Win32_UserProfile | Where-Object { $_.LocalPath -eq "C:\Users\utente" }
$profile.Delete()
```

