$global:DefaultConfig = @{
    'TrainingName' = 'Default Training';
    'MainDCPath' = "DC=domain,DC=local";
    'MainOUName' = "Test";
    'UsersOUName' = "Users";
    'GroupsOUName' = "Groups";
    'MailboxesOUName' = "Mailbox";
    'UsersCSVFileName' = "users.csv";
    'GroupsCSVFileName' = "groups.csv";
    'DefaultPassword' = "Qwerty12345";
    'DefaultEnabled' = $False;
    'GroupsForAll'="U¿ytkownicy";
    'DeletionProtected'=$False
}

$global:DefaultConfig | Export-Clixml ".\tmp.xml"
$global:CurrentConfig =  Import-Clixml ".\tmp.xml"
Remove-Item ".\tmp.xml" -Force
$global:cwd = Split-Path $MyInvocation.MyCommand.Definition
$global:Array = @{
    'MainWorkingDir' = Split-Path $global:cwd;
    'ConfigXMLFileName' = "mainconf.xml";
    'CurrentConfigXMLFileName' = "curconf.xml";
    'ConfigXMLFilePath' = "";
    'Language' = "eng";
    'BackupConfigFileName' = "backupconf.xml";
    'IncludeDirName' = "include";
    'TrainingDirName' = "trainings";
    'LanguageFileDefaultExtension' = ".ps1";
    'IncludeDirPath' = "";
    'CurrentLanguageFile' = "";
    'Version' = "0.1 alpha";
    'ProgramName' = "AD Helper";
    'UsersCSVContent' = "Title;GivenName;Surname;Organization;DisplayName;EmailAddress;telephoneNumber;SamAccountName";
    'GroupsCSVContent' = "Name;DisplayName;Alias";
    'Domain' = (gwmi Win32_ComputerSystem).domain.ToString()
}

$global:Array.IncludeDirPath = Join-Path -Path $global:Array.MainWorkingDir -ChildPath $global:Array.IncludeDirName
$global:Array.TrainingDirPath = Join-Path -Path $global:Array.MainWorkingDir -ChildPath $global:Array.TrainingDirName
$global:Array.CurrentLanguageFile = Join-Path -Path $global:Array.IncludeDirPath -ChildPath ($global:Array.Language + $global:Array.LanguageFileDefaultExtension)
$global:Array.ConfigXMLFilePath = Join-Path -Path $global:Array.IncludeDirPath -ChildPath $global:Array.ConfigXMLFileName

Function Array-ExportToXML {
    $global:Array | Export-Clixml (Join-Path -Path $global:Array.IncludeDirPath -ChildPath "mainconf.xml")
    $global:CurrentConfig | Export-Clixml (Join-Path -Path $global:Array.IncludeDirPath -ChildPath "curconf.xml")
}

Function Array-ImportFromXML {
    $global:Array = Import-Clixml (Join-Path -Path $global:Array.IncludeDirPath -ChildPath "mainconf.xml")
    If(!(Test-Path $global:Array.IncludeDirPath)) { 
        $global:Array.IncludeDirPath = $global:cwd
    }
    $global:CurrentConfig = Import-Clixml (Join-Path -Path $global:Array.IncludeDirPath -ChildPath "curconf.xml")
    If(!(Test-Path $global:Array.TrainingDirPath)) {
        $global:Array.TrainingDirPath = Join-Path -Path (Split-Path -Path $global:cwd) -ChildPath "trainings"
    }
}

If(Test-Path $global:Array.CurrentlanguageFile) {
    Write-Host "`n`tLanguage file loaded!" -ForegroundColor DarkGreen
} else {
    Write-Host "`n`tError!" -ForegroundColor Red -NoNewline
    Write-Host " Can't load language file." -ForegroundColor DarkRed
    exit
}

If(Test-Path $global:Array.ConfigXMLFilePath) {
    If((Get-ChildItem $global:Array.ConfigXMLFilePath | Where-Object { $_.Length -gt 1000 }) -eq $Null) {
        Write-Host "`n`tCritical ERROR!" -ForegroundColor Red -NoNewline
        Write-Host " Main Config File is corrupted!.`n" -ForegroundColor DarkRed
        exit
    } else {
        Array-ImportFromXML
        $global:Array.CurrentLanguageFile = Join-Path -Path $Array.IncludeDirPath -ChildPath ($Array.Language + $Array.LanguageFileDefaultExtension)
        Write-Host "`n`tConfig file" $global:Array.ConfigXMLFileName "loaded!" -ForegroundColor DarkGreen
        Start-Sleep 1
    }
} else {
    Write-Host "`n`tCan't locate" $global:Array.ConfigXMLFilePath -ForegroundColor Red
    Array-ExportToXML
    Write-Host "`n`tRestoring Default Configuration." -ForegroundColor DarkRed
    Start-Sleep 1
}

If(!(Test-Path (Join-Path -Path $global:Array.IncludeDirPath -ChildPath $global:Array.CurrentConfigXMLFileName))) {
    $global:DefaultConfig | Export-Clixml ".\tmp.xml"
    $global:CurrentConfig =  Import-Clixml ".\tmp.xml"
    Remove-Item ".\tmp.xml" -Force
    Array-ExportToXML
}
    
Array-ExportToXML
