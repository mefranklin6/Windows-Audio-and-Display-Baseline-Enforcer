Import-Module DisplayConfig
Import-Clixml C:\ProgramData\CTS\display_config_profile.xml | Use-DisplayConfig -UpdateAdapterId
