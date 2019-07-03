set-executionpolicy undefined -scope process -force
set-executionpolicy unrestricted -scope current -force

function install-registryKey($path, $name, $value, $type) {
    if(!(test-path $path)) {
        new-item -path $path -force | out-null
    }
     
    try {
        remove-itemProperty -LiteralPath $path -name $name -ErrorAction SilentlyContinue | out-null
    } catch {
    
    }
    
    new-itemProperty -LiteralPath $path -name $name -value $value -propertyType $type -force -ErrorAction SilentlyContinue | out-null
}

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
reg import .\RegistrySettings\RemoveAllUserFolders.reg
reg import .\RegistrySettings\SaveChrome.reg
reg import .\RegistrySettings\SetDriveIcons.reg

write-host "Adding Open With Registry Settings"
# The command paths must be hard coded and can't use environment variables so we have to add these this way
install-registryKey 'HKCU:\Software\Classes\*\shell\GVim' '(default)' 'Open with GVim' 'String'
install-registryKey 'HKCU:\Software\Classes\*\shell\GVim' 'Icon' '%userprofile%\settings\icons\vim.ico' 'String'
install-registryKey 'HKCU:\Software\Classes\*\shell\GVim\command' '(default)' '$env:userprofile\vim\vim80\gvim.exe "%1"' 'String'
install-registryKey 'HKCU:\Software\Classes\*\shell\Notepad' '(default)' 'Open with Notepad' 'String'
install-registryKey 'HKCU:\Software\Classes\*\shell\Notepad\command' '(default)' 'c:\windows\system32\notepad.exe "%1"' 'String'


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




