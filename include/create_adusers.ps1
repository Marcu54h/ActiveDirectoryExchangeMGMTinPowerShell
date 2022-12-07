# @Author: kpr Adam Marzec

# Set current working location to script-directory
Set-Location -Path (Split-Path $myinvocation.mycommand.definition) 

# Select language
$language = "eng"
$language = ".\include\" + $language + ".ps1"
. $language

Clear-Host
#Check if ActiveDirectory module is available
$adModuleAvailable = $False
(Get-Module -ListAvailable | Select-Object Name) | foreach {
    if($_.Name -eq "ActiveDirectory") { $adModuleAvailable = $True }
}
if(!$adModuleAvailable) {
    Write-Host $Speak.noADModule -ForegroundColor DarkRed
    exit
} else { Write-Host $Speak.adModuleOK -ForegroundColor Green }
try { Import-Module ActiveDirectory } catch { Write-Host $Speak.errorLoadingADModule -ForegroundColor Red }

# Make sure that encoding is ibm852
$CodePage = ($OutputEncoding | Select-Object CodePage).CodePage
if($CodePage -eq [Int32]852) {
    Write-Host $Speak.codePageOK -ForegroundColor Green
    Write-Host $Speak.fileEncodingInformation -ForegroundColor DarkGreen
} else {
    chcp 852 | Out-Null #change default console codepage to ibm852
    $OutputEncoding = [Console]::OutputEncoding
    $CodePage = ($OutputEncoding | Select-Object CodePage).CodePage
    Write-Host $Speak.codePageChanged $CodePage -ForegroundColor DarkGreen
    Write-Host $Speak.fileEncodingInformation -ForegroundColor DarkGreen
}

#Check if script is running on Exchange server
$EX = ((([Net.DNS]::GetHostByName($env:COMPUTERNAME)).AddressList | Select-Object IPAddressToString).IPAddressToString).EndsWith(".3")

#Load config file and set global variables
$configXML = ".\config.xml"  # Default name and location of configuration file
$Config = $Null              # Prepare variable for current configuration
$grCSV = "groups.csv"
$usCSV = "users.csv"
if(Test-Path $configXML) { # <--------------------------------------------------------------------------------|
    $Config = Import-Clixml $configXML
} else {  # If no config file is available these settings will be provided and saved to new $ConfigXML file
    Write-Host $Speak.configFileNotFound1 $configXML $Speak.configFileNotFound2 $configXML $Speak.configFileNotFound3
    $Config = @{
        'dcPath' = "DC=12dzzsd,DC=pmn,DC=int"
        'defaultPassword' = "Qwerty12345"
        'groupsForAll' = ("DRUKARKI_TAC","12DZ_TAC_SHAREPOINT")
        'defaultEnabled' = $False
        'mainOUName' = "TEST"
        'usersOUName' = "Users"
        'groupsOUName' = "Groups"
        'mailboxesOUName' = "Mailbox"
        'groupsCSV' = $grCSV
        'usersCSV' = $usCSV
        'trainingName' = ""
        'scriptLanguage' = $language
    } # END of $Config
    $Config | Export-Clixml $configXML
} #END of if else <-------------------------------------------------------------------------------------------|

$global:mnPath = "OU=" + $Config.mainOUName + "," + $Config.dcPath
$global:usPath = "OU=" + $Config.usersOUName + "," + $global:mnPath
$global:grPath = "OU=" + $Config.groupsOUName + "," + $global:mnPath
$global:mbPath = "OU=" + $Config.mailboxesOUName + "," + $global:mnPath
$global:defaultPassword = ConvertTo-SecureString $Config.defaultPassword -AsPlainText -Force
$global:groupsForAll = $Config.groupsForAll       # List of groups which all users will be members of
$global:defaultEnabled = $Config.defaultEnabled   # Default Enabled value for new-created accounts ($False/$True)
$global:groupsCSV = $Config.groupsCSV             # File with groups data
$global:usersCSV = $Config.usersCSV               # File with users data

$global:domain = (gwmi Win32_ComputerSystem).domain.ToString()
exit ###############################
if($EX) {
    Write-Host $Speak.exBannerTop -ForegroundColor Red
    Write-Host $Speak.exBannerContent -ForegroundColor White
    Write-Host $Speak.exBannerBottom -ForegroundColor Red
    Write-Host $Speak.createADStructQ -ForegroundColor Yellow
    $runOnEx = ($Host.UI.RawUI.ReadKey("noecho,includekeydown")).Character.ToString().tolower()
    if(!($runOnEx -like 'y')) { exit }
}
# Checking if Organizational Units provided by $global:Config are available in ActiveDirectory
# if not, these Organizationa Units will be created
if((Get-ADOrganizationalUnit -Filter * | Where-Object { $_.DistinguishedName -like $global:mnPath }).DistinguishedName -eq $Null) {
    Write-Host $Speak.mainOUCreate $global:mnPath -ForegroundColor Yellow
    New-ADOrganizationalUnit -Name $Config.mainOUName -Path $Config.dcPath
}
if((Get-ADOrganizationalUnit -Filter * | Where-Object { $_.DistinguishedName -like $global:usPath }).DistinguishedName -eq $Null) {
    Write-Host $Speak.ouCreate $global:usPath -ForegroundColor Yellow
    New-ADOrganizationalUnit -Name $Config.usersOUName -Path $global:mnPath
}
if((Get-ADOrganizationalUnit -Filter * | Where-Object { $_.DistinguishedName -like $global:grPath }).DistinguishedName -eq $Null) {
    Write-Host $Speak.ouCreate $global:grPath -ForegroundColor Yellow
    New-ADOrganizationalUnit -Name $Config.groupsOUName -Path $global:mnPath
}
if((Get-ADOrganizationalUnit -Filter * | Where-Object { $_.DistinguishedName -like $global:mbPath }).DistinguishedName -eq $Null) {
    Write-Host $Speak.ouCreate $global:mbPath -ForegroundColor Yellow
    New-ADOrganizationalUnit -Name $Config.mailboxesOUName -Path $global:mnPath
}

Write-Host "`n"
    # * * * * * * * * * * * * * * * * * * * * GROUPS CREATION * * * * * * * * * * * * * * * * * * * *
    try {
        #Import-Csv $global:groupsCSV -Delimiter ";" | ForEach-Object {
        Get-Content $global:groupsCSV -Encoding "String" | ConvertFrom-Csv -Delimiter ";" | ForEach-Object {
                $distName = "CN=" + $_.Name + "," + $global:grPath
                Write-Host $Speak.groupExist1 $_.Name $Speak.groupExist2 -ForegroundColor Blue
                if(dsquery group -samid $_.Name) {
                    Write-Host $Speak.found $_.Name -ForegroundColor Blue
                   ## LOCATION CHECKING
                    Write-Host $Speak.chGroupLocation -ForegroundColor Blue
                    if((Get-ADGroup $_.Name | Select-Object DistinguishedName).DistinguishedName -ne $distName) {
                        Write-Host $Speak.moving $_.Name $Speak.toNewLocation $global:grPath  -ForegroundColor Blue -NoNewline
                        $groupGUID = (Get-ADGroup $_.Name | Select-Object ObjectGUID).ObjectGUID
                        $groupGUID = $groupGUID.Guid
                        Move-ADObject -Identity $groupGUID -TargetPath $global:grPath
                        Write-Host $Speak.OK -ForegroundColor Green
                    }
                } else {
                    Write-Host $Speak.group $_.Name $Speak.notFoundCreating -NoNewline  -ForegroundColor Blue
                    $groupParams = @{
                        Name=$_.Name;
                        SamAccountName=$_.Name;
                        GroupCategory="Security";
                        GroupScope="Universal";
                        DisplayName=$_.DisplayName;
                        Path=$global:grPath;
                        OtherAttributes=@{'mail'=$_.Alias}
                    }
                    New-ADGroup @groupParams
                    Write-Host $Speak.group $_.Name $Speak.created -ForegroundColor Blue
                }
        }
    } catch { Write-Host "`tError loading" $global:groupsCSV "file!" -ForegroundColor Red }
    #END GROUPS CREATION

    # Get names of all AD 
    $adGroups = (Get-ADGroup -Filter * -Properties mail| Select-Object Name,DistinguishedName,mail)
    
    # Helpful Global variables
    $globalDisplayName = $Null
    $globalSamAccountName = $Null
    $newUser = $Null

    # * * * * * * * * * * * * * * * * * * * * USERS CREATION * * * * * * * * * * * * * * * * * * * *
#    try {
        #Import-Csv $global:usersCSV -Delimiter ";" | ForEach-Object {
        Get-Content $global:usersCSV -Encoding "String" | ConvertFrom-Csv -Delimiter ";" | ForEach-Object { # to process csv-file with polish chars
                $name = $_.Surname + " " + $_.GivenName
                $distName = "CN=" + $name + "," + $global:usPath
                Write-Host $Speak.chIfUser $name $Speak.exist -ForegroundColor Blue # IF START
                if(dsquery user -samid $_.SamAccountName) { # <=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-|
                    Write-Host "`tFound" $name -ForegroundColor Blue
                    ## LOCATION CHECKING
                    Write-Host $Speak.chIf $name $Speak.isInDeclaredLoc -ForegroundColor Blue -NoNewline
                    if((Get-ADUser $_.SamAccountName | Select-Object DistinguishedName).DistinguishedName -ne $distName) {
                        Write-Host $Speak.moving $name $Speak.toNewLocation $global:usPath  -ForegroundColor Blue -NoNewline
                        $userGUID = (Get-ADUser $_.SamAccountName | Select-Object ObjectGUID).ObjectGUID
                        $userGUID = $userGUID.Guid
                        Move-ADObject -Identity $userGUID -TargetPath $global:usPath
                        Write-Host " OK!`n" -ForegroundColor Green
                    } else { Write-Host " OK!`n" -ForegroundColor Green } 
                         #                                           IF END                                          ELSE START
                } else { # <=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=|
                    Write-Host $Speak.user $name $Speak.notFoundCreate -ForegroundColor Blue -NoNewline
                    $globalDisplayName = $_.DisplayName
                    $globalSamAccountName = $_.SamAccountName
                    $params = @{  # <-----------------------------------------------------------------|
                        SamAccountName=$_.SamAccountName;
                        GivenName=$_.GivenName;
                        Surname=$_.Surname;
                        Name=$_.Surname + " " + $_.GivenName;
                        Path=$global:usPath;
                        DisplayName=$_.DisplayName;
                        Title=$_.Title;
                        Organization=$_.Organization;
                        AccountPassword= $global:defaultPassword;
                        UserPrincipalName=$_.SamAccountName + "@" + $domain;
                        Enabled = $global:defaultEnabled;
                        OtherAttributes=@{'telephoneNumber'=$_.telephoneNumber;'mail'=$_.EmailAddress}
                    }  # <----------------------------------------------------------------------------|
                    #try {
                        if($params.OtherAttributes.mail -like "") { $params.OtherAttributes.mail = "default@mail.local" }
                        $userMail = $params.OtherAttributes.mail
                        $newUser = New-ADUser @params -PassThru # PassThru returns new-created user to variable
                        Write-Host " OK!`n" -ForegroundColor Green
                    #} catch { Write-Host "Something went wrong while creating new AD user!" }
                    # Checking user/group membership.
                    # $adGroups contains groups (Name,DistinguishedName,mail) located in whole ActiveDirectory
                    $adGroups | foreach { # <===================================================================================| 
                        $n = $_.Name
                        if(($globalDisplayName -like $n + "*") -and (($globalDisplayName.length - $n.length) -lt 4 )) { # <-|
                            $groupDN = $_.DistinguishedName
                            $groupMail = $_.mail
                            if($newUser){  # <------------------------------------------|
                                Add-ADGroupMember -Identity $groupDN -Members $newUser
                                Write-Host "`tUser" $newUser.Name "added to the group" $groupDN "`n" -ForegroundColor Blue
                                if(!($userMail -like $groupMail) -and ($userMail -like "default@mail.local")) {
                                    Write-host `t"No e-mail found... setting" $groupMail "for" $newUser.Name -ForegroundColor DarkRed
                                    Set-ADUser $newUser.SamAccountName -Clear 'mail'
                                    Set-ADUser $newUser.SamAccountName -Add @{'mail'=$groupMail}
                                }
                            }  else { Write-Host $Speak.userEmpty -ForegroundColor Red } # <--------------------|
                        }  # <----------------------------------------------------------------------------------------------|
                    }  # <=======================================================================================================|
                    if($newUser) { # <-------------------------------------------|
                        $global:groupsForAll | foreach { # <-----------------|
                            Add-ADGroupMember -Identity $_ -Members $newUser
                            Write-Host $Speak.user $newUser.Name $Speak.addedToGroup $_ "`n" -ForegroundColor Blue
                        }  # <-----------------------------------------------|
                    } else { Write-Host "`tEmpty user!" -ForegroundColor Red } # <----------------------|                    ELSE END
                } # <=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=|
        }
#    } catch {
 #       Write-Host "Error loading AD_users.csv file!"
#    }
    # * * * * * * * * * * * * * * * * * * * *END USERS CREATION * * * * * * * * * * * * * * * * * * * *

# * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 
#
# ========     Script below this line will run only on EX server ==============================================================================
#
# * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 * 8 *
#Write-Host "`n`tWaiting fo AD to synchronize...`n" -ForegroundColor White 
#Start-Sleep 10

if($EX) {
    Write-Host $Speak.exBannerTop -ForegroundColor Red
    Write-Host $Speak.exBannerContent -ForegroundColor White
    Write-Host $Speak.exBannerBottom -ForegroundColor Red
    # Create Shared mailboxes
    $mailBoxName = Get-Mailbox | Select-Object Name
    Get-Content $global:groupsCSV -Encoding "String" | ConvertFrom-Csv -Delimiter ";" | ForEach-Object {
        $tmpName = $_.Name
        $tmpDN = $_.DisplayName
        $tmpAlias = $_.Alias
        $Name = $tmpName + "_MB"
        # Check if mailbox $Name already exists
        if(($mailBoxName | Where-Object { $_.Name -like $Name }).Name -like $Name ) {
            Write-Host $Speak.mailbox $Name $Speak.alreadyExists -ForegroundColor DarkRed
        } else {
            Write-Host $Speak.newSharedMB $tmpName "..." -ForegroundColor Cyan -NoNewline
            New-Mailbox -Shared -Name $Name -DisplayName $tmpDN -Alias $tmpAlias | Out-Null
            Write-Host $Speak.mailbox $Name $Speak.created -ForegroundColor Green
            Start-Sleep 1
            $mailBoxName = Get-Mailbox | Select-Object Name
        }

        Write-Host $Speak.addGroupPerm $tmpName $Speak.toMailbox $Name "..." -ForegroundColor Cyan -NoNewline
        $groupDN = (Get-ADGroup $tmpName | Select-Object DistinguishedName).DistinguishedName
        Add-MailboxPermission -Identity $Name -User $groupDN -AccessRights "FullAccess" | Out-Null
        Start-Sleep 1
        Get-Mailbox $Name | Add-ADPermission -User $groupDN -ExtendedRights "Send-As" | Out-Null
        Start-Sleep 1
        #Add-MailboxPermission -Identity $Name -User $groupDN -AccessRights "SendAs"
        Write-Host $Speak.OK2 -ForegroundColor Green
        Start-Sleep 1
        $mailboxDN = (Get-ADUser $Name | Select-Object DistinguishedName).DistinguishedName
        Write-Host $Speak.moving $Name $Speak.to $global:mbPath "..." -ForegroundColor Cyan -NoNewline
        Move-ADObject -Identity $mailboxDN -TargetPath $global:mbPath
        Write-Host $Speak.OK2 -ForegroundColor Green
    }
}
Write-Host $Speak.credentials -ForegroundColor Green -NoNewline
Write-Host "`t`tB" -ForegroundColor Red -NoNewline
Write-Host "y" -ForegroundColor Yellow -NoNewline
Write-Host "e" -ForegroundColor Green -NoNewline
Write-Host "! " -ForegroundColor Blue -NoNewline
Write-Host "B" -ForegroundColor Red -NoNewline
Write-Host "y" -ForegroundColor Yellow -NoNewline
Write-Host "e" -ForegroundColor Green -NoNewline
Write-Host "!`n" -ForegroundColor Blue