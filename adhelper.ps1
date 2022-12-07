Set-Location (Split-Path $MyInvocation.MyCommand.Definition)

. ".\include\chstr.ps1"
$exchange = $False
If(Get-AppDir -Mask "Microsoft Exchange Server*") { $exchange = $True }

$admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

If($admin) {
    . ".\include\adutils.ps1"
    
    $bgColor = $Host.UI.RawUI.BackgroundColor
    $Host.UI.RawUI.BackgroundColor = "Black"

    $mainMenuChoice = ""
    $trMenuChoice = ""
    $optMenuChoice = ""
    $freshStart = $True

    # MAIN LOOP
    Function Switch-ToTrainingMenu {
        while(!($trMenuChoice -like "b"))
            {
                Draw-TrainingMenu
                
                $trMenuChoice = ($Host.UI.RawUI.ReadKey("noecho,includekeydown")).Character.ToString().ToLower()
                
                If($trMenuChoice -like "x") {
                    Execute-ADStructPrep
                    $Host.UI.RawUI.ReadKey("noecho,includekeydown") # pause - wait for key
                }
                
                If($trMenuChoice -like "u") {
                    Open-UsersData
                }
                
                If($trMenuChoice -like "g") {
                    Open-GroupsData
                }
                
                If($trMenuChoice -like "k") {
                    Unlock-LockedADUsers
                    Start-Sleep 2
                }
            }
    }

    Function Switch-ToOptionsMenu {
        while(!($optMenuChoice -like "b"))
        {
            Draw-OptionMenu
            
            $optMenuChoice = ($Host.UI.RawUI.ReadKey("noecho,includekeydown")).Character.ToString().ToLower()
            
            If($optMenuChoice -like "l") {
                Write-Host $Speak.currentLanguageIs -ForegroundColor DarkGray -NoNewline
                Write-Host "" $global:Array.Language -ForegroundColor Gray
                Write-Host $Speak.availableLanguages -ForegroundColor DarkGray
                Change-Language
            }
            
            If($optMenuChoice -like "p") {
                Change-DefaultPassword
                Start-Sleep 2
            }
        }
    }

    while(!($mainMenuChoice -like "q"))
    {
        If(!($global:CurrentConfig.TrainingName -like $global:DefaultConfig.TrainingName) -and $freshStart) {
            Switch-ToTrainingMenu
            $freshStart = $False
        }
        
        Draw-MainMenu
        $mainMenuChoice = ($Host.UI.RawUI.ReadKey("noecho,includekeydown")).Character.ToString().ToLower()
        If($mainMenuChoice -like "l") {
            Get-Trainings
            $Host.UI.RawUI.ReadKey("noecho,includekeydown")
        }
        
        If(!($global:CurrentConfig.TrainingName -like $global:DefaultConfig.TrainingName) -and ($mainMenuChoice -like "t")) {
            Switch-ToTrainingMenu
        }
        
        If($mainMenuChoice -like "n") {
            New-Training
            Start-Sleep 2
        }
        
        If($mainMenuChoice -like "s") {
            Get-Trainings
            Set-CurrentTraining
            Start-Sleep 2
        }
        
        If($mainMenuChoice -like "r") {
            Get-Trainings
            Remove-Training
            Start-Sleep 2
        }
        
        If($mainMenuChoice -like 'o') {
            Switch-ToOptionsMenu
        }
    }
    Bye-Bye
    $Host.UI.RawUI.BackgroundColor = $bgColor
    Clear-Host

} else {
    Write-Host $Speak.noAdminError -ForegroundColor DarkRed
    $fullScriptPath = $MyInvocation.MyCommand.Definition
    $cwd = Split-Path $fullScriptPath
    $args = ""
    If($exchange) {
        $args = "-Command . 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto; Set-Location '$cwd'; & '$fullScriptPath'" 
    } else {
        $args = " -Command Set-Location " + "'$cwd'; & '$fullScriptPath'"
    }
    try {
        Start-Process PowerShell -ArgumentList $args -Verb "runAs"
    } catch {
        Write-Host $Speak.noAuthorization -ForegroundColor Red
        Start-Sleep 3
    }
}