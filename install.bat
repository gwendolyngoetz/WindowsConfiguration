@echo on

rem Importing Registry Settings
reg import .\RegistrySettings\AquaSnapSettings.reg
reg import .\RegistrySettings\CmdAutoRunAlias.reg
reg import .\RegistrySettings\NoLockScreen.reg
reg import .\RegistrySettings\OpenWithSettings.reg
reg import .\RegistrySettings\RemoveAllUserFolders.reg
reg import .\RegistrySettings\ResetFolderRedirection.reg
reg import .\RegistrySettings\SaveChrome.reg
reg import .\RegistrySettings\SetDriveIcons.reg

rem Set Persistant Environment Variables
setx src d:\src
setx home %userprofile%
setx ConEmuPromptNames NO