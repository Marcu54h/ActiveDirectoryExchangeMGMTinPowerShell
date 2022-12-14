Function Get-AppDir {
    Param (
        [Parameter(Mandatory=$True)]
        [String]$Mask
    )
    if([IntPtr]::Size -eq 4) {  # check if system is 32 bit
        $regPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    } else {
        $regPath = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
    }
    Get-ItemProperty $regPath | .{process{if($_.DisplayName -and $_.UninstallString) { $_ } }} | Select InstallLocation,DisplayName | Where-Object { !($_.InstallLocation -like "") } | Where-Object { $_.DisplayName -like $Mask } | FL
    #$result = Get-InstalledApps | Where { $_.DisplayName -like $appToMatch }
}

Function Get-EditorExecutable {
    Param(
        [Parameter(Mandatory=$False)]
        [String]$Name
    )
    If($Name -like "notepad++") {
        IF(Test-Path "C:\Program Files\Notepad++\notepad++.exe") { return "C:\Program Files\Notepad++\notepad++.exe" }
        If(Test-Path "C:\Program Files (x86)\Notepad++\notepad++.exe") { return "C:\Program Files (x86)\Notepad++\notepad++.exe" }
    }
    If($Name -like "notepad") {
        If(Test-Path "C:\Windows\notepad.exe") { return "C:\Windows\notepad.exe" }
    }
    If($Name -like "excel") {
        If(Test-Path "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE") { return "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" }
        If(Test-Path "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE") { return "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" }
    }
    If(Test-Path "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE") { return "C:\Program Files (x86)\Microsoft Office\Office14\EXCEL.EXE" }
    If(Test-Path "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE") { return "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE" }
    IF(Test-Path "C:\Program Files\Notepad++\notepad++.exe") { return "C:\Program Files\Notepad++\notepad++.exe" }
    If(Test-Path "C:\Program Files (x86)\Notepad++\notepad++.exe") { return "C:\Program Files (x86)\Notepad++\notepad++.exe" }
    If(Test-Path "C:\Windows\notepad.exe") { return "C:\Windows\notepad.exe" }
    If(Test-Path "C:\Windows\System32\notepad.exe") { return "C:\Windows\System32\notepad.exe" }
}

Function Change-WindowSize {
    Param(
        [Parameter(Mandatory=$True)]
        [Int]$Width,
        [Parameter(Mandatory=$True)]
        [Int]$Height,
        [Parameter(Mandatory=$False)]
        [String]$WindowTitle
    )
    $maxWidth = (Get-Host).UI.RawUI.MaxPhysicalWindowSize.Width
    #Write-Host "Max Width is" $maxWidth
    $maxHeight = (Get-Host).UI.RawUI.MaxPhysicalWindowSize.Height
    #Write-Host "Max Height is" $maxHeight
    If($Height -lt 42) { $Height = 42 }
    If($Height -gt $maxHeight) { $Height = $maxHeight }
    If($Width -lt 60) { $Width = 60 }
    If($Width -gt $maxWidth) { $Width = $maxWidth }
    If(!($WindowTitle -eq $Null)) {
        (Get-Host).UI.RawUI.WindowTitle = $WindowTitle 
    }
    $wndSize = (Get-Host).UI.RawUI.WindowSize
    $wndSize.Width = $Width
    $wndSize.Height = $Height
    $buffSize = (Get-Host).UI.RawUI.BufferSize
    $buffSize.Width = $Width
    $buffSize.Height = $Height
    (Get-Host).UI.RawUI.BufferSize = $buffSize
    (Get-Host).UI.RawUI.WindowSize = $wndSize
}