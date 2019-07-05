param(
  [switch] $skipInstall
)

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
    'powershell-core',
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


if ($skipInstall -eq $false) {
    write-host "Install Packages"
    foreach ($package in $packages) {
        choco install $package --limit-output --yes
    }
}


write-host "Importing Registry Settings"
reg import .\RegistrySettings\AquaSnapSettings.reg
reg import .\RegistrySettings\CleanupExplorerContextMenus.reg
reg import .\RegistrySettings\CmdAutoRunAlias.reg
reg import .\RegistrySettings\NoLockScreen.reg
reg import .\RegistrySettings\RemoveAllUserFolders.reg
reg import .\RegistrySettings\SaveChrome.reg
reg import .\RegistrySettings\SetDriveIcons.reg


write-host "Adding Open With Registry Settings"
# The command paths must be hard coded and can't use environment variables so we have to add these this way
new-item -force -ErrorAction SilentlyContinue 'HKCU:\Software\Classes\*\shell\GVim' | out-null
new-item -force -ErrorAction SilentlyContinue 'HKCU:\Software\Classes\*\shell\GVim\command' | out-null
new-item -force -ErrorAction SilentlyContinue 'HKCU:\Software\Classes\*\shell\Notepad' | out-null
new-item -force -ErrorAction SilentlyContinue 'HKCU:\Software\Classes\*\shell\Notepad\command' | out-null

new-itemProperty -force -ErrorAction SilentlyContinue -propertyType 'String' -LiteralPath 'HKCU:\Software\Classes\*\shell\GVim'             -name '(default)'  -value 'Open with GVim' | out-null
new-itemProperty -force -ErrorAction SilentlyContinue -propertyType 'String' -LiteralPath 'HKCU:\Software\Classes\*\shell\GVim'             -name 'Icon'       -value '%userprofile%\settings\icons\vim.ico' | out-null
new-itemProperty -force -ErrorAction SilentlyContinue -propertyType 'String' -LiteralPath 'HKCU:\Software\Classes\*\shell\GVim\command'     -name '(default)'  -value '$env:userprofile\vim\vim80\gvim.exe "%1"' | out-null
new-itemProperty -force -ErrorAction SilentlyContinue -propertyType 'String' -LiteralPath 'HKCU:\Software\Classes\*\shell\Notepad'          -name '(default)'  -value 'Open with Notepad' | out-null
new-itemProperty -force -ErrorAction SilentlyContinue -propertyType 'String' -LiteralPath 'HKCU:\Software\Classes\*\shell\Notepad\command'  -name '(default)'  -value 'c:\windows\system32\notepad.exe "%1"' | out-null


write-host "Setting variales"
$srcDir="d:\src"
$settingsDir="$env:userprofile\settings"
$iconsDir="$settingsDir\icons"


write-host "Adding custom environment variables"
setx src $srcDir | out-null
setx home $env:userprofile | out-null
setx ConEmuPromptNames NO | out-null


write-host "Creating Directories"
try { md -force $srcDir | out-null } catch { write-host $_.exception.message }
try { md -force $settingsDir | out-null } catch { write-host $_.exception.message }
try { md -force $iconsDir | out-null } catch { write-host $_.exception.message }
try { md -force "$env:userprofile\.vim" | out-null } catch { write-host $_.exception.message }
try { md -force "$env:userprofile\.tmp" | out-null } catch { write-host $_.exception.message }


write-host "Copy Files"
cp .\alias.cmd $env:userprofile\alias.cmd
cp .\icons\* $env:userprofile\settings\icons
cp .\ConEmu\ConEmu.xml $env:appdata


write-host "DONE"

