Function Check-ADModule {
    $adModuleAvailable = $False
    (Get-Module -ListAvailable | Select-Object Name) | foreach {
        if($_.Name -eq "ActiveDirectory") { $adModuleAvailable = $True }
    }
    if(!$adModuleAvailable) {
        Write-Host $Speak.noADModule -ForegroundColor DarkRed
        exit
    } else { Write-Host $Speak.adModuleOK -ForegroundColor Green }
    try {
        Import-Module ActiveDirectory
    } catch {
        Write-Host $Speak.errorLoadingADModule -ForegroundColor Red
    }
}
Check-ADModule

Function Check-CodePage {
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
}
Check-CodePage

Function Check-ConnectionWithDC {
    $dcPath = $global:CurrentConfig.MainDCPath
    $pattern = 'DC=(\w{1,})?\b'
    $dcURL = ([RegEx]::Matches($dcPath, $pattern) | ForEach-Object {$_.Value})-join"."-replace'DC=',''
    If(Test-Connection -ComputerName $dcURL -Quiet -Count 1) {
        Write-Host $Speak.dcResolveOK $dcURL $Speak.OK -ForegroundColor DarkGreen
        return $True
    } else {
        Write-Host $Speak.cantResolveDC1 $dcURL $Speak.cantResolveDC2 -ForegroundColor Red
        return $False
    }
}

Function Create-UserStructure {
    $curTrDir = Join-Path -Path $global:Array.TrainingDirPath -ChildPath $global:CurrentConfig.TrainingName
    $usrCSV = Join-Path -Path $curTrDir -ChildPath $global:CurrentConfig.UsersCSVFileName
    Get-Content $usrCSV -Encoding "String" | ConvertFrom-Csv -Delimiter ";" | ForEach-Object { # to process csv-file with polish chars
        $name = $_.Surname + " " + $_.GivenName
        $usersOUDN = "OU=" + $global:CurrentConfig.UsersOUName + ",OU=" + $global:CurrentConfig.MainOUName + "," + $global:CurrentConfig.MainDCPath
        $distName = "CN=" + $name + "," + $usersOUDN
        Write-Host $Speak.chIfUser $name $Speak.exist -ForegroundColor Blue
                                                                                 # IF START
        If(dsquery user -samid $_.SamAccountName) { # <=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-|
            Write-Host "`tFound" $name -ForegroundColor Blue
            ## LOCATION CHECKING
            Write-Host $Speak.chIf $name $Speak.isInDeclaredLoc -ForegroundColor Blue -NoNewline
            
            If((Get-ADUser $_.SamAccountName | Select-Object DistinguishedName).DistinguishedName -ne $distName) {
                Write-Host $Speak.moving $name $Speak.toNewLocation $usersOUDN  -ForegroundColor Blue -NoNewline
                $userGUID = (Get-ADUser $_.SamAccountName | Select-Object ObjectGUID).ObjectGUID
                $userGUID = $userGUID.Guid
                Move-ADObject -Identity $userGUID -TargetPath $usersOUDN
                Write-Host " OK!`n" -ForegroundColor Green
            } else {
                Write-Host " OK!`n" -ForegroundColor Green
            } 
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
                Path=$usersOUDN;
                DisplayName=$_.DisplayName;
                Title=$_.Title;
                Organization=$_.Organization;
                AccountPassword= ConvertTo-SecureString -String $global:CurrentConfig.DefaultPassword -AsPlainText -Force;
                UserPrincipalName=$_.SamAccountName + "@" + $global:Array.Domain;
                Enabled = $global:CurrentConfig.DefaultEnabled;
                OtherAttributes=@{'telephoneNumber'=$_.telephoneNumber;'mail'=$_.EmailAddress}
            }  # <----------------------------------------------------------------------------|
            #try {
            If($params.OtherAttributes.mail -like "") {
                $params.OtherAttributes.mail = "default@mail.local"
            }
            $userMail = $params.OtherAttributes.mail
            $newUser = New-ADUser @params -PassThru # PassThru returns new-created user to variable
            Write-Host " OK!`n" -ForegroundColor Green
            #} catch { Write-Host "Something went wrong while creating new AD user!" }
            # Checking user/group membership.
            # $adGroups contains groups (Name,DistinguishedName,mail) located in whole ActiveDirectory
            $adGroups = (Get-ADGroup -Filter * -Properties mail| Select-Object Name,DistinguishedName,mail)
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
                $global:CurrentConfig.GroupsForAll | foreach { # <-----------------|
                    Add-ADGroupMember -Identity $_ -Members $newUser
                    Write-Host $Speak.user $newUser.Name $Speak.addedToGroup $_ "`n" -ForegroundColor Blue
                }  # <-----------------------------------------------|
            } else { Write-Host "`tEmpty user!" -ForegroundColor Red } # <----------------------|                    ELSE END
        } # <=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=--=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=|
    }
}

Function Create-GroupStructure {
    $curTrDir = Join-Path -Path $global:Array.TrainingDirPath -ChildPath $global:CurrentConfig.TrainingName
    $groupCSV = Join-Path -Path $curTrDir -ChildPath $global:CurrentConfig.GroupsCSVFileName
    #try {
        Get-Content $groupCSV -Encoding "String" | ConvertFrom-Csv -Delimiter ";" | ForEach-Object {
            $groupsOUDN = "OU=" + $global:CurrentConfig.GroupsOUName + ",OU=" + $global:CurrentConfig.MainOUName + "," + $global:CurrentConfig.MainDCPath
            Write-Host "`t" $groupsOUDN
            $distName = "CN=" + $_.Name + "," + $groupsOUDN
            Write-Host $Speak.groupExist1 $_.Name $Speak.groupExist2 -ForegroundColor Blue
            if(dsquery group -samid $_.Name) {
                Write-Host $Speak.found $_.Name -ForegroundColor Blue
               ## LOCATION CHECKING
                Write-Host $Speak.chGroupLocation -ForegroundColor Blue
                if((Get-ADGroup $_.Name | Select-Object DistinguishedName).DistinguishedName -ne $distName) {
                    Write-Host $Speak.moving $_.Name $Speak.toNewLocation $groupsOUDN  -ForegroundColor Blue -NoNewline
                    $groupGUID = (Get-ADGroup $_.Name | Select-Object ObjectGUID).ObjectGUID
                    $groupGUID = $groupGUID.Guid
                    Move-ADObject -Identity $groupGUID -TargetPath $groupsOUDN
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
                    Path=$groupsOUDN;
                    OtherAttributes=@{'mail'=($_.Alias + "@" + $global:Array.Domain)}
                }
                New-ADGroup @groupParams
                Write-Host $Speak.group $_.Name $Speak.created -ForegroundColor Blue
            }
        }
   # } catch { Write-Host "`tError loading" $global:CurrentConfig.GroupsCSVFileName "file!" -ForegroundColor Red }
}

Function Create-Mailboxes {
    Write-Host $Speak.exBannerTop -ForegroundColor Red
    Write-Host $Speak.exBannerContent -ForegroundColor White
    Write-Host $Speak.exBannerBottom -ForegroundColor Red
    # Create Shared mailboxes
    $mailBoxName = Get-Mailbox | Select-Object Name
    $curTrDir = Join-Path -Path $global:Array.TrainingDirPath -ChildPath $global:CurrentConfig.TrainingName
    $groupCSV = Join-Path -Path $curTrDir -ChildPath $global:CurrentConfig.GroupsCSVFileName
    Get-Content $groupCSV -Encoding "String" | ConvertFrom-Csv -Delimiter ";" | ForEach-Object {
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
        $mailboxesOUDN = "OU=" + $global:CurrentConfig.MailboxesOUName + ",OU=" + $global:CurrentConfig.MainOUName + "," + $global:CurrentConfig.MainDCPath
        $mailboxDN = (Get-ADUser $Name | Select-Object DistinguishedName).DistinguishedName
        Write-Host $Speak.moving $Name $Speak.to $mailboxesOUDN "..." -ForegroundColor Cyan -NoNewline
        Move-ADObject -Identity $mailboxDN -TargetPath $mailboxesOUDN
        Write-Host $Speak.OK2 -ForegroundColor Green
    }
}

Function Create-OUStructure {
    $mainOUDN = "OU=" + $global:CurrentConfig.MainOUName + "," + $global:CurrentConfig.MainDCPath
    $usersOUDN = "OU=" + $global:CurrentConfig.UsersOUName + "," + $mainOUDN
    $groupsOUDN = "OU=" + $global:CurrentConfig.GroupsOUName + "," + $mainOUDN
    $mailboxesOUDN = "OU=" + $global:CurrentConfig.MailboxesOUName + "," + $mainOUDN
    
    if((Get-ADOrganizationalUnit -Filter * | Where-Object { $_.DistinguishedName -like $mainOUDN }).DistinguishedName -eq $Null) {
        Write-Host $Speak.mainOUCreate $mainOUDN -ForegroundColor Yellow
        New-ADOrganizationalUnit -Name $global:CurrentConfig.MainOUName -Path $global:CurrentConfig.MainDCPath -ProtectedFromAccidentalDeletion $global:CurrentConfig.DeletionProtected
    }
    
    if((Get-ADOrganizationalUnit -Filter * | Where-Object { $_.DistinguishedName -like $usersOUDN }).DistinguishedName -eq $Null) {
        Write-Host $Speak.ouCreate $usersOUDN -ForegroundColor Yellow
        New-ADOrganizationalUnit -Name $global:CurrentConfig.UsersOUName -Path $mainOUDN -ProtectedFromAccidentalDeletion $global:CurrentConfig.DeletionProtected
    }
    
    if((Get-ADOrganizationalUnit -Filter * | Where-Object { $_.DistinguishedName -like $groupsOUDN }).DistinguishedName -eq $Null) {
        Write-Host $Speak.ouCreate $groupsOUDN -ForegroundColor Yellow
        New-ADOrganizationalUnit -Name $global:CurrentConfig.GroupsOUName -Path $mainOUDN -ProtectedFromAccidentalDeletion $global:CurrentConfig.DeletionProtected
    }
    
    if((Get-ADOrganizationalUnit -Filter * | Where-Object { $_.DistinguishedName -like $mailboxesOUDN }).DistinguishedName -eq $Null) {
        Write-Host $Speak.ouCreate $mailboxesOUDN -ForegroundColor Yellow
        New-ADOrganizationalUnit -Name $global:CurrentConfig.MailboxesOUName -Path $mainOUDN -ProtectedFromAccidentalDeletion $global:CurrentConfig.DeletionProtected
    }
}

Function Unlock-LockedAdUsers {
    $searchBase = "OU=" + $global:CurrentConfig.UsersOUName + ",OU=" + $global:CurrentConfig.MainOUName + "," + $global:CurrentConfig.MainDCPath
    If(Check-ConnectionWithDC) {
        try {
            If(Search-ADAccount -LockedOut -SearchBase $searchBase -ErrorAction Stop) {
                Search-ADAccount -LockedOut -SearchBase $searchBase -ErrorAction Stop | Unlock-ADAccount
                Write-Host $Speak.inOU $searchBase -ForegroundColor Cyan
                Write-Host $Speak.accountsUnlocked -ForegroundColor Yellow
            } else {
                Write-Host $Speak.noLockedAccounts -ForegroundColor DarkGreen
            }
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            Write-Host $Speak.errorOUNotFound -ForegroundColor DarkRed
        }
    }
}

Function Execute-ADStructPrep {
    If(Check-ConnectionWithDC) {
        Create-OUStructure
        Create-GroupStructure
        Create-UserStructure
        If(Get-AppDir -Mask "Microsoft Exchange Server*") {
            Create-Mailboxes
        }
        Write-Host $Speak.done -ForegroundColor Green
    }
}