$Speak = @{
    'noADModule' = "`n`tActive-Directory module is not available. It's imposible to proceed. Bye!`n";
    'adModuleOK' = "`n`tActiveDirectory module is availabale. We may proceed.`n";
    'codePageOK' = "`tCodepage OK! (ibm852)`n";
    'fileEncodingInformation' = "`tPlease make sure that encoding of files which contain users`n`tand groups data are also encoded as Windows 1250 (imb852) files.`n`tOtherwise '�' will appear as 'Ł' and so on.`n";
    'errorLoadingADModule' = "`tError while loading module: ActiveDirectory!";
    'codePageChanged' = "`tCodepage changed to imb";
    'configFileNotFound1' = "`tConfiguration file";
    'configFileNotFound2' = "not found!`n`tCreating";
    'configFileNotFound3' = "with default settings...`n";
    'exBannerTop' = "`n`t-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-";
    'exBannerContent' = "`t`t!!! Script is running on EXCHANGE SERVER !!!";
    'exBannerBottom' = "`t-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-*-_-`n";
    'createADStructQ' = "`n`tDo you want to continue to create whole AD structure on`n`t`tExchange Server? [y/n]`n";
    'mainOUCreate' = "`tCreating main OU";
    'ouCreate' = "`tCreating OU";
    'groupExist1' = "`tChecking if group";
    'groupExist2' = "exists...";
    'found' = "`tFound";
    'chGroupLocation' = "`tChecking if group is in declared location...`n";
    'moving' = "`tMoving";
    'toNewLocation' = "to new location...";
    'OK' = " OK!`n";
    'OK2' = "OK!";
    'group' = "`tGroup";
    'notFoundCreating' = "not found.`n`tCreating...";
    'created' = "created.`n";
    'credentials' = "`n`nCreated by: Adam Marzec`ne-mail: adam.marzec@gmail.com`n12 bdow Szczecin";
    'to' = 'to';
    'mailbox' = "`tMailbox";
    'alreadyExists' = "already exists!";
    'newSharedMB' = "`n`tCreating new Shared Mailbox for";
    'userEmpty' = "`tUser empty!";
    'user' = "`tUser";
    'addedToGroup' = "added to the group";
    'notFoundCreate' = "not found.`n`tCreating...";
    'chIf' = "`tChecking if";
    'isInDeclaredLoc' = "is in declared location...";
    'chIfUser' = "`tChecking if user";
    'exist' = "exists...";
    'addGroupPerm' = "`tAdding permissions for group";
    'toMailbox' = "to mailbox";
    'trainingNameNotSet' = "`n`tTraining name is empty.";
    'newTrainingQuestion' = "`n`tDo you want to create new one? [Y/N]";
    'enterTrainingName' = "Enter name for new training";
    'createTrainingStructure' = "`n`tStructure for new training doesn't exist. Do you want to create it? [Y/N]";
    'availableLanguages' = "`n`tAvailable languages`n`tEnter 'eng' for ENGLISH'`n`tEnter 'pol' for POLISH";
    'wrongLanguage' = "`n`tWrong language entered. Leaving...";
    'noAdminError' = "`n`tError! No administrator priviledges occured.";
    'noAuthorization' = "`n`tNo authorization!";
    'noLockedAccounts' = "`n`tNo Locked Accouts found!";
    # ============= chstr.ps1 ===================
    'trainingAlreadyExist' = "`n`tGiven training already exist. Aborting.";
    'trainingCreated' = "`n`tTraining successfully created.";
    'removeTrainingError' = "`n`tChoosen training doesn't exist! Available trainings are:";
    'training' = "`n`tTraining";
    'trainingActivated' = "has been activated!";
    'trainingRemoved' = "has been removed!";
    'trainingConfFileNotExist' = "`n`tCurrent training configuration file doesn't exist! Creating...";
    'qSetTrainingData' = "`t`nDo you want to set basic settings for current training? [Y/N]";
    'promptForOUNamesChange' = "`n`tOU names are set to default. Do you want to change them? [Y/N]";
    'enterMainOU' = "`n`tPlease enter name for main OU";
    'enterUsersOU' = "`n`tPlease enter name for Users OU: ";
    'enterMailboxesOU' = "`n`tPlease enter name for Mailboxes OU: ";
    'enterGroupsOU' = "`n`tPlease enter name for Groups OU: ";
    'exclamationConfigSwap' = "`n`tYou're trying to change Current Training to Training with the same Name.`n`tDo you want to swam Current Config with BackupConfig?";
    'infConfigsSwaped' = "`n`tConfigurations where swaped!.";
    'mainDCPathNotSet' = "`n`tMain Domain Controler Path is set to default Values.`n`tDo you want to change it now? [Y/N]";
    'enterMainDCPath' = "`n`tPlease enter path to main DC (ex. 'DC=domain,DC=local')";
    'currentPassword' = "`n`tCurrent password is";
    'qChangePassword' = "`n`tDo you want to change current password? [Y/N]";
    'currentPasswordChanged' = "`n`tCurrent password has been changed.";
    'passwordNotChanged' = "`n`tPassword hasn't been changed.";
    # ============= MAIN MENU ========================
    'mainMenuTitle' = "`t`t" + $global:Array.ProgramName + " " + $global:Array.Version + "`t`t`t";
    'switchToTrainingMenu' = "`t[T] Enter Training MENU`t`t`t`t";
    'switchToOptionMenu' = "`t[O] Options`t`t`t`t`t";
    'trMenuExecuteADStruct' = "`t[X] Execute AD and Mailbox creation process`t"
    'unlockLockedADUsers' = "`t[K] Unlock locked AD users`t`t`t";
    'mainMenuCurrentTraining' = "`tCurrent training is";
    'mainMenuOpenUsersData' = "`t[U] Open Users Data file`t`t`t";
    'mainMenuOpenGroupsData' = "`t[G] Open Groups Data file`t`t`t";
    'mainMenuListTrainings' = "`t[L] List all trainings`t`t`t`t";
    'mainMenuNewTraining' = "`t[N] Create New Training`t`t`t`t";
    'mainMenuSetCurrentTraining' = "`t[S] Set Current Training`t`t`t";
    'mainMenuRemoveTraining' = "`t[R] Remove Training`t`t`t`t";
    'mainMenuLine' = "`t`t`t`t`t`t`t";
    'mainMenuQuit' = "`t[Q] Quit`t`t`t`t`t";
    'backToMainMenu' = "`t[B] Back to MAIN MENU`t`t`t`t";
    # ============= OPTION MENU ======================
    'optionMenuTitle' = "`t`tOPTIONS MENU`t`t`t`t";
    'optionMenuLanguage' = "`t[L] Change display language`t`t`t";
    'currentLanguageIs' = "`n`tCurrent language is";
    'optionMenuPassword' = "`t[P] Change default password`t`t`t";
    # ============= adutils ==========================
    'cantResolveDC1' = "`n`tCurrent Domain Controler";
    'cantResolveDC2' = "can't be resolved!";
    'dcResolveOK' = "`n`tResolving";
    'accountsUnlocked' = "`n`tAll locked users accounts has been unlocked.";
    'inOU' = "`n`tIn organizational unit";
    'done' = "`n`tDone!";
    'errorOUNotFound' = "`n`tCan't find given OU!"
}