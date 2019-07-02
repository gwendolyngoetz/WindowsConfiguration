set-executionpolicy undefined -scope process -force
set-executionpolicy unrestricted -scope current -force

# Install applications from Chocolatey
if ((test-path C:\ProgramData\chocolatey\bin\choco.exe) -eq $false) {
    write-host "Installing Chocolatey" 
    iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1')) 
} 

$packages = @(
    '7zip', 
    'autohotkey.portable',
    'azurestorageexplorer',
    'cmake',
    'conemu',
    'curl',
    'elevate.native',
    'everything',
    'far',
    'fiddler', 
    'filezilla', 
    'firefox',
    'git',
    'gitextensions',
    'googlechrome',
    'html-tidy',
    'ilspy',
    'kdiff3',
    'keepass',
    'lockhunter',
    'meld',
    'mremoteng',
    'nimbletext',
    'nodejs',
    'notepad2',
    'notepadplusplus',
    'nuget.commandline',
    'nugetpackageexplorer',
    'oldcalc',
    'postman',
    'powershell',
    'putty',
    'reshack',
    'rktools.2003',
    'shexview',
    'shmnview',
    'smtp4dev',
    'sumatrapdf',
    'superorca',
    'sysinternals',
    'vim',
    'virtualbox',
    'virtualclonedrive',
    'visualstudiocode',
    'wget',
    'windirstat',
    'winmerge2011.portable',
    'winscp',
    'xmlstarlet',
    'yeoman'
)


write-host "Install Packages"
foreach ($package in $packages) {
    choco install $package --limit-output --yes
}


write-host "Importing Registry Settings"
reg import .\RegistrySettings\AquaSnapSettings.reg
reg import .\RegistrySettings\CmdAutoRunAlias.reg
reg import .\RegistrySettings\NoLockScreen.reg
reg import .\RegistrySettings\OpenWithSettings.reg
reg import .\RegistrySettings\RemoveAllUserFolders.reg
reg import .\RegistrySettings\ResetFolderRedirection.reg
reg import .\RegistrySettings\SaveChrome.reg
reg import .\RegistrySettings\SetDriveIcons.reg


write-host "Adding custom environment variables"
setx src $srcDir
setx home $env:userprofile
setx ConEmuPromptNames NO


write-host "Creating Directories"
$srcDir="d:\src"
$settingsDir="$env:userprofile\settings"
$iconsDir="$settingsDir\icons"

try { md -force $srcDir } catch { write-host $_.exception.message }
try { md -force $settingsDir } catch { write-host $_.exception.message }
try { md -force $iconsDir } catch { write-host $_.exception.message }
try { md -force "$env:userprofile\.vim" } catch { write-host $_.exception.message }
try { md -force "$env:userprofile\.tmp" } catch { write-host $_.exception.message }


write-host "Copy Files"
cp .\alias.cmd $env:userprofile\alias.cmd
cp .\icons\* $env:userprofile\settings\icons

write-host "DONE"




