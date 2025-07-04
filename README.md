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
```
```powershell
xcopy /Y ultravnc.ini "C:\Program Files\uvnc bvba\UltraVNC\"
```
```powershell
net stop uvnc_service
```
```powershell
net start uvnc_service
```

Silent uninstallation of UltraVNC:
```powershell
"C:\Program Files\uvnc bvba\UltraVNC\unins000.exe" /silent
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

Write-Host "Inizio script di installazione font"

# Scarica i font dal percorso condiviso al temp
try {
    Copy-Item -Path $fontFolder\*.ttf -Destination $tempFolder -Force
    Write-Host "Font copiati con successo"
} catch {
    Write-Host "Errore durante la copia dei font: $_"
    Write-EventLog -LogName Application -Source "FontInstallScript" -EventID 1000 -EntryType Error -Message "Errore durante la copia dei font: $_"
    Stop-Transcript
    exit
}

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
        try {
            Copy-Item -Path $fontPath -Destination $destinationPath -Force
            $regPath = "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"
            Set-ItemProperty -Path $regPath -Name $fontName -Value $fontName
            Write-Host "Installazione completata per: $fontName"
            Write-EventLog -LogName Application -Source "FontInstallScript" -EventID 1002 -EntryType Information -Message "Installazione completata per: $fontName"
        } catch {
            Write-Host "Errore durante l'installazione di $fontName: $_"
            Write-EventLog -LogName Application -Source "FontInstallScript" -EventID 1000 -EntryType Error -Message "Errore durante l'installazione di $fontName: $_"
        }
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

Get all enabled computers in the domain:
```powershell
# Ottieni solo i computer abilitati nel dominio
Get-ADComputer -Filter {Enabled -eq $true} | Select-Object Name, DNSHostName, OperatingSystem
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

### Remote Management
Configure WinRM and enable PowerShell Remoting:
```powershell
# Configura WinRM con impostazioni di base
winrm quickconfig -q
```

# Configura il servizio per l'avvio automatico e lo avvia
```powershell
Set-Service WinRM -StartupType Automatic
```
```powershell
Start-Service WinRM
```

# Abilita PowerShell Remoting
```powershell
Enable-PSRemoting -Force
```

# Configura il listener WinRM
```powershell
winrm create winrm/config/listener?Address=*+Transport=HTTP
```

# Configura le impostazioni di sicurezza di base
```powershell
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force
```

# Apri le porte nel firewall di Windows
```powershell
New-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)" -Name "Windows Remote Management (HTTP-In)" -Profile Any -LocalPort 5985 -Protocol TCP
```

# Verifica lo stato del servizio WinRM
```powershell
Get-Service WinRM
```

Connettersi a un computer remoto:
# Connessione base a un computer remoto
```powershell
Enter-PSSession -ComputerName "NomeComputer"
```

# Connessione usando un account specifico
```powershell
Enter-PSSession -ComputerName "NomeComputer" -Credential "DOMINIO\Utente"
```

# Eseguire un comando su un computer remoto senza aprire una sessione
```powershell
Invoke-Command -ComputerName "NomeComputer" -ScriptBlock { Get-Service } -Credential (Get-Credential)
```
# Connessione a più computer contemporaneamente
```powershell
Invoke-Command -ComputerName "PC1", "PC2", "PC3" -ScriptBlock { Get-Service } -Credential (Get-Credential)
```

### Autologon (Accesso Automatico)

Imposta l'autologin con Sysinternals Autologon.exe:
```powershell
Autologon.exe username domain password /accepteula
```

> ⚠️ **Nota di sicurezza:**  
> L'uso di password in chiaro nei comandi o negli script non è sicuro.  
> Per un autologin più sicuro, imposta le credenziali tramite l'interfaccia di Autologon.exe (esegui senza parametri) e inserisci le informazioni manualmente:  
> 
> 1. Avvia `Autologon.exe` come amministratore.
> 2. Inserisci username, dominio e password.
> 3. Clicca su "Enable".
> 
> In questo modo la password viene salvata in modo cifrato nel registro di sistema.

