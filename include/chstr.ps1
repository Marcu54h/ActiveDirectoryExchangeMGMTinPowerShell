If(Test-Path ".\include\array.ps1") { . ".\include\array.ps1" } else { Write-Host "`nCan't locate 'array.ps1' file in 'include' dir!`n" -ForegroundColor Red }
. $global:Array.CurrentLanguageFile
. (Join-Path -Path $global:Array.IncludeDirPath -ChildPath "utils.ps1")

$wndWidth = 110
$wndHeight = 41
$wndTitle = $global:Array.ProgramName + " " + $global:Array.Version

Function Get-Trainings {
    Array-ImportFromXML
    Get-ChildItem $global:Array.TrainingDirPath | Select-Object Mode,Name | Where-Object { $_.Mode -eq "d----" }
}

Function New-Training {
    #[CmdLetBinding()]
    Param (
        [Parameter(Mandatory=$True)]
        [String]$Name
    )
    #Begin {}
    #Process {}
    #End{}
    if((Get-Trainings | Where-Object { $_.Name -like $Name }) -eq $Null) {
        $trDir = (Join-Path -Path $global:Array.TrainingDirPath -ChildPath $Name)
        New-Item -Path $trDir -ItemType Directory | Out-Null
        $global:DefaultConfig | Export-Clixml ".\tmp.xml"
        $global:CurrentConfig =  Import-Clixml ".\tmp.xml"
        Remove-Item ".\tmp.xml" -Force
        $global:CurrentConfig.TrainingName = $Name
        New-Item -Path (Join-Path -Path $trDir -ChildPath $global:CurrentConfig.UsersCSVFileName) -ItemType File | Out-Null
        Add-Content -Value $global:Array.UsersCSVContent -Path (Join-Path -Path $trDir -ChildPath $global:CurrentConfig.UsersCSVFileName)
        New-Item -Path (Join-Path -Path $trDir -ChildPath $global:CurrentConfig.GroupsCSVFileName) -ItemType File | Out-Null
        Add-Content -Value $global:Array.GroupsCSVContent -Path (Join-Path -Path $trDir -ChildPath $global:CurrentConfig.GroupsCSVFileName)
        Write-Host $Speak.createTrainingStructure -ForegroundColor Cyan
        $ouChange = $Host.UI.RawUI.ReadKey("noecho,includekeydown").Character.ToString().ToLower()
        If($ouChange -like 'y') {
            $tmp = Read-Host -Prompt $Speak.enterMainOU
            If(!($tmp -like "")) {
                $global:CurrentConfig.MainOUName = $tmp
            }
            
            $tmp = Read-Host -Prompt $Speak.enterUsersOU
            If(!($tmp -like "")) {
                $global:CurrentConfig.UsersOUName = $tmp
            }
            
            $tmp = Read-Host -Prompt $Speak.enterGroupsOU
            If(!($tmp -like "")) {
                $global:CurrentConfig.GroupsOUName = $tmp
            }
            
            $tmp = Read-Host -Prompt $Speak.enterMailboxesOU
            If(!($tmp -like "")) {
                $global:CurrentConfig.MailboxesOUName = $tmp
            }
            
            $tmp = Read-Host -Prompt $Speak.enterMainDCPath
            If(!($tmp -like "")) {
                $global:CurrentConfig.MainDCPath = $tmp
            }  
        }
        If($global:CurrentConfig.MainDCPath -like $global:DefaultConfig.MainDCPath) {
            Write-Host $Speak.mainDCPathNotSet -ForegroundColor Cyan -NoNewline
            $tmp = $Host.UI.RawUI.ReadKey("noecho,includekeydown").Character.ToString().ToLower()
            If($tmp -like "y") {
                $tmp = Read-Host -Prompt $Speak.enterMainDCPath
            }
        }
        
        $global:CurrentConfig | Export-Clixml (Join-Path -Path $trDir -ChildPath $global:Array.BackupConfigFileName)
        Array-ExportToXML
        Write-Host $Speak.trainingCreated -ForegroundColor Green
    } else {
        Write-Host $Speak.trainingAlreadyExist -ForegroundColor DarkRed
    }
}

Function Remove-Training {
    Param (
        [Parameter(Mandatory=$True)]
        [String]$Name
    )
    if(Get-Trainings | Where-Object { $_.Name -like $Name }) {
        Remove-Item (Join-Path -Path $global:Array.TrainingDirPath -ChildPath $Name) -Recurse -Force
        Write-Host $Speak.training $Name $Speak.trainingRemoved -ForegroundColor Yellow
        $global:CurrentConfig = $global:DefaultConfig
        Array-SaveToXML $global:Array.ConfigXMLFilePath
    } else {
        Write-Host $Speak.removeTrainingError -ForegroundColor DarkRed
        Get-Trainings
    }
}

Function Set-CurrentTraining {
    Param (
        [Parameter(Mandatory=$True)]
        [String]$Name
    )
    if(Get-Trainings | Where-Object { $_.Name -like $Name }) {
        $bckUpFile = Join-Path -Path (Join-Path -Path $global:Array.TrainingDirPath -ChildPath $Name) -ChildPath $global:Array.BackupConfigFileName
        If($Name -like $global:CurrentConfig.TrainingName) {
            Write-Host $Speak.exclamationConfigSwap -ForegroundColor Cyan
            $q = Read-Host -Prompt "`n`tYour answer: [Y/N]" 
            If($q -like "y" -or $q -like "Y") {
                $global:CurrentConfig = Import-Clixml $bckUpFile
                Write-Host $Speak.infConfigsSwaped -ForegroundColor Cyan
                Array-SaveToXML $global:Array.ConfigXMLFilePath
            }
        } else {
            $global:CurrentConfig = Import-Clixml $bckUpFile
            Array-ExportToXML
        }
        Write-Host $Speak.training $Name $Speak.trainingActivated -ForegroundColor Cyan
    }
}

Function Open-UsersData {
    $uCSV = $global:Array.TrainingDirPath + "\" + $global:CurrentConfig.TrainingName
    Start-Process (Get-EditorExecutable -Name "excel") ('"' + $uCSV + "\" +  $global:CurrentConfig.UsersCSVFileName + '"')
}
Function Open-GroupsData {
    $gCSV = Join-Path -Path $global:Array.TrainingDirPath -ChildPath $global:CurrentConfig.TrainingName
    Start-Process (Get-EditorExecutable -Name "excel") ('"' + $gCSV + "\" + $global:CurrentConfig.GroupsCSVFileName + '"')
}

Function Change-Language {
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Language
    )
    If($Language -like "eng" -or $Language -like "pol" -or $Language -like "ENG" -or $Language -like "POL") {
       $global:Array.Language = $Language
       Array-ExportToXML
       $global:Array.CurrentLanguageFile = Join-Path -Path $Array.IncludeDirPath -ChildPath ($Array.Language + $Array.LanguageFileDefaultExtension)
       . $global:Array.CurrentLanguageFile
       Start-Sleep 1
       Draw-OptionMenu 
    } else {
        Write-Host $Speak.wrongLanguage -ForegroundColor DarkRed
        Start-Sleep 1
    }
}

Function Change-DefaultPassword {
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Password
    )
    If($global:CurrentConfig.TrainingName -ne $global:DefaultConfig) {
        Write-Host $Speak.currentPassword $global:CurrentConfig.DefaultPassword -ForegroundColor Cyan
        Write-Host $Speak.qChangePassword -ForegroundColor Cyan
        $q = ($Host.UI.RawUI.ReadKey("noecho,includekeydown")).Character.ToString().ToLower()
        If($q -like "y") {
            $trDir = (Join-Path -Path $global:Array.TrainingDirPath -ChildPath $global:CurrentConfig.TrainingName)
            $global:CurrentConfig.DefaultPassword = $Password
            $global:CurrentConfig | Export-Clixml (Join-Path -Path $trDir -ChildPath $global:Array.BackupConfigFileName)
            Array-ExportToXML
            Write-Host $Speak.currentPasswordChanged -ForegroundColor Magenta
        } else {
            Write-Host $Speak.passwordNotChanged -ForegroundColor DarkMagenta
        }
    }
}

Function Draw-MainMenu {
    Change-WindowSize -Height $wndHwight -Width $wndWidth -WindowTitle ($wndTitle + " - MAIN MENU")
    Clear-Host
    #Write-Host "Current Training" $global:Array.CurrentConfig.TrainingName
    #Write-Host "Last Training" $global:Array.LastConfig.TrainingName
    #Write-Host "Default Training" $global:Array.DefaultConfig.TrainingName
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkMagenta
    Write-Host $Speak.mainMenuTitle -BackgroundColor DarkMagenta -ForegroundColor White
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkMagenta
    Write-Host $Speak.mainMenuCurrentTraining $global:CurrentConfig.TrainingName -ForegroundColor Green
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkGreen
    Write-Host $Speak.switchToTrainingMenu -BackgroundColor DarkGreen -ForegroundColor Red
    
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkGreen
    Write-Host $Speak.mainMenuListTrainings -BackgroundColor DarkGreen -ForegroundColor Yellow
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkGreen
    Write-Host $Speak.mainMenuNewTraining -BackgroundColor DarkGreen -ForegroundColor Yellow
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkGreen
    Write-Host $Speak.mainMenuSetCurrentTraining -BackgroundColor DarkGreen -ForegroundColor Yellow
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkGreen
    Write-Host $Speak.mainMenuRemoveTraining -BackgroundColor DarkGreen -ForegroundColor Yellow
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkGreen
    Write-Host $Speak.switchToOptionMenu -BackgroundColor DarkGreen -ForegroundColor Red
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkGreen
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkRed
    Write-Host $Speak.mainMenuQuit -BackgroundColor DarkRed -ForegroundColor Black
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkRed
}

Function Draw-TrainingMenu {
    Change-WindowSize -Height $wndHwight -Width $wndWidth -WindowTitle ($wndTitle + " - TRAINING MENU")
    Clear-Host
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkMagenta
    Write-Host $Speak.mainMenuTitle -BackgroundColor DarkMagenta -ForegroundColor White
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkMagenta
    Write-Host $Speak.mainMenuCurrentTraining $global:CurrentConfig.TrainingName -ForegroundColor Green
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkBlue
    
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkBlue
    Write-Host $Speak.trMenuExecuteADStruct -BackgroundColor DarkBlue -ForegroundColor Yellow
    
    
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkBlue
    Write-Host $Speak.mainMenuOpenUsersData -BackgroundColor DarkBlue -ForegroundColor Yellow
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkBlue
    Write-Host $Speak.mainMenuOpenGroupsData -BackgroundColor DarkBlue -ForegroundColor Yellow
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkBlue
    Write-Host $Speak.unlockLockedADUsers -BackgroundColor DarkBlue -ForegroundColor Yellow
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkBlue
    
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkRed
    Write-Host $Speak.backToMainMenu -BackgroundColor DarkRed -ForegroundColor Black
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkRed
}

Function Draw-OptionMenu {
    Change-WindowSize -Height $wndHwight -Width $wndWidth -WindowTitle ($wndTitle + " - OPTIONS MENU")
    Clear-Host
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkGreen
    Write-Host $Speak.optionMenuTitle -BackgroundColor DarkGreen -ForegroundColor White
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkGreen
    
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkMagenta
    Write-Host $Speak.optionMenuPassword -BackgroundColor DarkMagenta -ForegroundColor Cyan
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkMagenta
    Write-Host $Speak.optionMenuLanguage -BackgroundColor DarkMagenta -ForegroundColor Cyan
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkMagenta
    
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkRed
    Write-Host $Speak.backToMainMenu -BackgroundColor DarkRed -ForegroundColor Black
    Write-Host $Speak.mainMenuLine -BackgroundColor DarkRed
}

Function Bye-Bye {
    Write-Host $Speak.credentials -ForegroundColor Green -NoNewline
    Write-Host "`t`t`tB" -ForegroundColor Red -NoNewline
    Write-Host "y" -ForegroundColor Yellow -NoNewline
    Write-Host "e" -ForegroundColor Green -NoNewline
    Write-Host "! " -ForegroundColor Blue -NoNewline
    Write-Host "B" -ForegroundColor Red -NoNewline
    Write-Host "y" -ForegroundColor Yellow -NoNewline
    Write-Host "e" -ForegroundColor Green -NoNewline
    Write-Host "!`n" -ForegroundColor Blue
    Start-Sleep 5
}