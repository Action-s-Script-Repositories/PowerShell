#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------

#region Helper Functions
function Get-ScriptDirectory
{
<#
	.SYNOPSIS
		Get-ScriptDirectory returns the proper location of the script.

	.OUTPUTS
		System.String
	
	.NOTES
		Returns the correct path within a packaged executable.
#>
	[OutputType([string])]
	param ()
	if ($null -ne $hostinvocation)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

function Paint-FocusBorder([System.Windows.Forms.Control]$control)
{
	# get the parent control (usually the form itself)
	$parent = $control.Parent
	$parent.Refresh()
	if ($control.Focused)
	{
		$control.BackColor = "LightBlue"
		$pen = [System.Drawing.Pen]::new('Red', 2)
	}
	else
	{
		$control.BackColor = "White"
		$pen = [System.Drawing.Pen]::new($parent.BackColor, 2)
	}
	$rect = [System.Drawing.Rectangle]::new($control.Location, $control.Size)
	$rect.Inflate(1, 1)
	$parent.CreateGraphics().DrawRectangle($pen, $rect)
}

function Load-Module ($m)
{
	# If module is imported say that and do nothing
	if (Get-Module | Where-Object { $_.Name -eq $m })
	{
		return 0
	}
	else
	{
		# If module is not imported, but available on disk then import
		if (Get-Module -ListAvailable | Where-Object { $_.Name -eq $m })
		{
			Import-Module $m -Verbose
			return 1
		}
		else
		{
			# If module is not imported, not available on disk, but is in online gallery then install and import
			if (Find-Module -Name $m | Where-Object { $_.Name -eq $m })
			{
				Install-Module -Name $m -Force -Verbose -Scope CurrentUser
				Import-Module $m -Verbose
				return 2
			}
			else
			{
				# If module is not imported, not available and not in online gallery then abort
				return 3
			}
		}
	}
}

function Set-UILanguage ($Culture)
{
	if ($Culture -like "nl-*")
	{
		#region Dutch language
		$Language = "Dutch"
		
		### General Form Controls
		$FormLabel = "Active Directory Gebruikers en Groepen v$($VersionNumber) - $($OrgName)"
		$SSTitle = 'Laden...'
		$SSUpdateLanguageList = 'Updaten taallijst'
		$SSSettingFormControls = 'Initialiseren besturingselementen'
		$SSSettingUILanguage = 'Taal instellen'
		$SSConfiguringRemotePSSession = 'PS-remoting sessie configureren'
		$SSRefreshingMembership = 'Groeplidmaatschap verversen...'
		$SSSearchingGroups = 'Groepen zoeken...'
		$SSSearchingUsers = 'Gebruikers zoeken...'
		
		$ButtonSearch = "&Zoeken"
		$ButtonAdd = "&Toevoegen"
		$ButtonRemove = "&Verwijderen"
		$ButtonAgain = "&Opnieuw"
		$ButtonCopy = "&Kopiëren"
		$ButtonExport = "&Exporteren"
		$ButtonExit = "&Afsluiten"
		$LabelVerboseOutput = 'Uitgebreide uitvoer:'
		$LabelLanguage = 'Taal:'
		$LabelDisplayName = 'Weergavenaam:'
		$GroupSearchControls = 'Zoekbediening'
		$TabPage0 = 'Gebruikers'
		$TabPage1 = 'Groepen'
		$MembersTabPage0 = 'Leden'
		$MembersTabPage1 = 'Lid van'
		
		### Form Labels
		#### Users Tab
		$LabelUsername = 'Gebruikersnaam:'
		$LabelGroupMembership = 'Groeplidmaatschappen:'
		$LabelDepartment = 'Afdeling:'
		$LabelFunction = 'Functie:'
		$LabelSamAccountName = 'Gebruikersnaam:'
		$LabelUserPrincipalName = 'User Principal Naam:'
		$LabelEmailAddress = 'Emailadres:'
		$GroupUserDetails = 'Gebruikersdetails'
		
		##### Add Group to User Form
		$AFFormName = 'Active Directory Gebruikers - Groep(en) toevoegen'
		$AFlblADGroup = 'AD groep:'
		$AFbtnAdd = "&Toevoegen"
		$AFbtnCancel = "&Annuleren"
		$AFbtnRetrieve = "&Ophalen"
		$AFNoADGroup = 'Geen groepsnaam ingegeven.'
		$AFADGroupFound = 'Groep(en) gevonden:'
		$AFADGroupInvalid = 'Geen groep(en) gevonden.'
		$AFGroupsFound = 'Gevonden groep(en):'
		$AFGroupsToBeAdded = 'Groep(en) toe te voegen:'
		
		#### Groups Tab
		$GRPLabelGroupname = 'Groepsnaam:'
		$GRPLabelGroupCategory = 'Groep categorie:'
		$GRPLabelGroupScope = 'Groepsbereik:'
		$GRPLabelGroupInfo = 'Groep info:'
		$GRPLabelManagedBy = 'Beheerd door:'
		$GRPLabelDirectMemberCount = 'Aantal directe leden:'
		$GRPLabelNestedGroupsCount = 'Aantal geneste groepen:'
		$GRPLabelMemberOfCount = 'Aantal lid van:'
		$GRPGroupGroupDetails = 'Groep details'
		
		##### Select Group Form
		$SGFormName = 'Active Directory Groepen - Groep Selecteren'
		$SGFormNameUsers = 'Active Directory Gebruikers - Gebruiker Selecteren'
		$SGbtnCancel = "&Annuleren"
		$SGbtnOK = "&OK"
		
		##### Add User to Group Form
		$AUFormName = 'Active Directory Groepen - Gebruiker(s) toevoegen'
		$AUlblADUser = 'AD gebruiker:'
		$AUNoADUser = 'Geen gebruikersnaam ingegeven.'
		$AUUserFound = 'Gebruiker(s) gevonden:'
		$AUADUserInvalid = 'Geen gebruiker(s) gevonden.'
		$AUUsersFound = 'Gevonden gebruiker(s):'
		$AUUsersToBeAdded = 'Gebruiker(s) toe te voegen:'
				
		### Tooltip text
		#### General Form Controls
		$TTcmbLanguage = 'Hier kan de formulier taal gekozen worden.'
		$TTbtnExit = 'Deze actie zal het formulier sluiten.'
		$TTrtbOutput = 'Dit is het log scherm.'
		
		#### Users Tab
		$TTtxtUsername = 'Vul hier de gebruikersnaam in van de gebruiker die je wilt opzoeken.'
		$TTbtnSearch = "Deze actie dient als eerste uitgevoerd te worden. Als de gebruiker gevonden wordt zullen de `'Toevoegen`', `'Verwijderen`', `'Kopiëren`' en `'Exporteren`' knoppen beschikbaar komen."
		$TTbtnAdd = 'Deze actie zal een pop-up geven waarin een Active Directory groep opgezocht kan worden waaraan de gebruiker toegevoegd zal worden.'
		$TTbtnRemove = 'Deze actie zal de gebruiker uit de geselecteerde groep(en) halen.'
		$TTbtnAgain = 'Deze actie zal het formulier terugzetten naar de beginstaat.'
		$TTbtnCopy = 'Deze actie kopieert de geselecteerde groep(en) naar het klembord.'
		$TTbtnExport = 'Deze actie exporteert de geselecteerde groep(en) naar een bestand.'
		$TTlstADGroups = 'Dit scherm zal de huidige groepen laten zien waarvan de opgegeven gebruiker lid is.'
		
		##### Add AD Group Form
		$TTAddtxtADGroup = 'Vul dit veld in met de AD-groep (of jokerteken) waarnaar u wilt zoeken.'
		$TTAddlstAddADGroups = "Deze keuzelijst bevat de gevonden groepen, gebaseerd op de zoekcriteria.`r`nDubbelklikken op een groep zal deze naar de andere keuzelijst sturen, zodat hij klaar is om aan de gebruiker toegevoegd te worden.`r`nCtrl+A selecteert alle gevonden groepen.`r`n Door op Ctrl+C te drukken, worden de geselecteerde groepen naar het klembord gekopieerd."
		$TTAddlstGroupsToBeAdded = "Deze keuzelijst bevat de groepen die klaar zijn om toegevoegd te worden aan de gespecificeerde gebruiker.`r`nDubbelklikken op een groep zal deze verwijderen uit de lijst.`r`nDoor op Ctrl+V te drukken terwijl de keuzelijst is geselecteerd, wordt de inhoud van het klembord geplakt."
		$TTAddbtnAdd = 'Deze actie loopt door de geselecteerde groepen en voegt ze toe aan de opgegeven gebruiker.'
		$TTAddbtnCancel = 'Deze actie zal het toevoegen annuleren en het formulier sluiten.'
		$TTAddbtnAgain = 'Deze actie zal het formulier terugzetten naar de oorspronkelijke staat.'
		$TTAddbtnRetrieve = 'Deze actie zoekt naar Active Directory-groepen op basis van jouw invoer.'
		$TTAddbtnLeft = "Deze actie zal de geselecteerde groep(en) verwijderen uit de lijst `'Groep(en) toe te voegen`'."
		$TTAddbtnRight = "Deze actie zal de geselecteerde groep(en) uit de `'Gevonden groep(en)`' lijst toevoegen aan de `'Groep(en) toe te voegen`' lijst."
		$TTAddrtbAddOutput = 'Dit is het log scherm.'
		
		##### Select Group Form
		$TTSGbtnCancel = 'Deze actie annuleert de groepsselectie.'
		$TTSGbtnOK = 'Deze actie zal de gemarkeerde groep gebruiken voor verdere acties.'
		$TTSGlstSelectObject = "Deze lijst bevat de groepen die zijn gevonden op basis van invoer.`r`n Selecteer een groep om verder te gaan."
		
		##### Add AD Group Form
		$TTAddUtxtADUser = 'Vul dit veld in met de AD-gebruiker (of jokerteken) waarnaar u wilt zoeken.'
		$TTAddUlstAddADUsers = "Deze keuzelijst bevat de gevonden gebruikers, gebaseerd op de zoekcriteria.`r`nDubbelklikken op een gebruiker zal deze naar de andere keuzelijst sturen, zodat hij klaar is om aan de gebruiker toegevoegd te worden.`r`nCtrl+A selecteert alle gevonden gebruikers.`r`n Door op Ctrl+C te drukken, worden de geselecteerde gebruikers naar het klembord gekopieerd."
		$TTAddUlstUsersToBeAdded = "Deze keuzelijst bevat de gebruikers die klaar zijn om toegevoegd te worden aan de gespecificeerde groep.`r`nDubbelklikken op een gebruiker zal deze verwijderen uit de lijst.`r`nDoor op Ctrl+V te drukken terwijl de keuzelijst is geselecteerd, wordt de inhoud van het klembord geplakt."
		$TTAddUbtnAdd = 'Deze actie loopt door de geselecteerde gebruikers en voegt ze toe aan de opgegeven groep.'
		$TTAddUbtnCancel = 'Deze actie zal het toevoegen annuleren en het formulier sluiten.'
		$TTAddUbtnAgain = 'Deze actie zal het formulier terugzetten naar de oorspronkelijke staat.'
		$TTAddUbtnRetrieve = 'Deze actie zoekt naar Active Directory-gebruikers op basis van jouw invoer.'
		$TTAddUbtnLeft = "Deze actie zal de geselecteerde gebruiker(s) verwijderen uit de lijst `'Gebruiker(s) toe te voegen`'."
		$TTAddUbtnRight = "Deze actie zal de geselecteerde gebruiker(s) uit de `'Gevonden gebruiker(s)`' lijst toevoegen aan de `'Gebruiker(s) toe te voegen`' lijst."
		$TTAddUrtbAddOutput = 'Dit is het log scherm.'
		
		#### Groups Tab
		$TTGRPtxtGroupname = 'Vul hier de groepsnaam in van de groep die je wilt opzoeken.'
		$TTGRPbtnSearch = "Deze actie dient als eerste uitgevoerd te worden. Als de groep gevonden wordt zullen de `'Toevoegen`', `'Verwijderen`', `'Kopiëren`' en `'Exporteren`' knoppen beschikbaar komen."
		$TTGRPbtnAdd = 'Deze actie zal een pop-up geven waarin een Active Directory user opgezocht kan worden die aan de groep toegevoegd zal worden.'
		$TTGRPbtnRemove = 'Deze actie zal de geselecteerde gebruiker(s) uit de groep halen.'
		$TTGRPbtnAgain = 'Deze actie zal het formulier terugzetten naar de beginstaat.'
		$TTGRPbtnCopy = 'Deze actie kopieert de geselecteerde gebruiker(s) naar het klembord.'
		$TTGRPbtnExport = 'Deze actie exporteert de geselecteerde gebruiker(s) naar een bestand.'
		$TTlstGroupMembers = 'Dit scherm zal de huidige leden laten zien van de opgegeven groep.'
		$TTlstGroupMemberOf = 'Dit scherm toont de huidige groepen waarvan de gespecificeerde groep lid is.'
		
		### Verbose output
		#### General strings
		$VBInfo = 'INFO:'
		$VBWarning = 'WAARSCHUWING:'
		$VBError = 'FOUT:'
		$VBSuccess = 'SUCCES:'
		$VBModuleCheck = 'Controleren op'
		$VBModule = 'module'
		$VBModuleAlreadyImported = 'module is al geïmporteerd.'
		$VBModuleAvailable = 'module beschikbaar. Importeren...'
		$VBModuleImported = 'module geïmporteerd.'
		$VBModuleSearch = 'Zoeken naar'
		$VBModuleOnline = 'module in online gallerij.'
		$VBModuleFound = 'module gevonden. Installeren...'
		$VBModuleInstalled = 'module geïnstalleerd. Importeren...'
		$VBModuleFailed = 'module niet geïmporteerd, niet beschikbaar en niet gevonden in de online gallerij.'
		$VBPSSessionSetup = 'PS-remoting sessie aan het opzetten.'
		$VBPSSessionSetupSuccess = 'PS-remoting verbinding tot stand gebracht.'
		$VBPSSessionSetupFailed = 'PS-remoting kan geen verbinding maken. Neem contact op met uw Systeembeheerder'
		$VBPSSessionImport = 'PS-remoting sessie importeren.'
		$VBPSSessionImportSuccess = 'PS-remoting sessie geïmporteerd.'
		$VBPSSessionImportFailed = 'PS-remoting kan sessie niet importeren. Neem contact op met uw Systeembeheerder.'
		$VBPSSessionInvalid = 'PS-remoting Geen geldige sessie gevonden. Neem contact op met uw Systeembeheerder'
		$VBButtonNotImplemented = 'Deze knop is nog niet geïmplementeerd.'
		$VBNoCBContent = 'Het klembord is leeg.'
		$VBCBPasted = 'Klembordinhoud geplakt.'
		$VBTooFewChar = "Te weining karakters ingegeven:`r`n`tGebruikersnaam: minimaal 2`r`n`tGroepsnaam: minimaal 4`r`n"
		
		#### Users Tab
		$VBNoUserName = 'Vul alsjeblieft een gebruikersnaam in waarop je wilt zoeken.'
		$VBNoUserFound = 'Geen gebruiker gevonden met de volgende gebruikersnaam'
		$VBNoGroup = 'Geen groep(en) geselecteerd. Selecteer 1 of meerdere groepen.'
		$VBUsersToBeDeleted = 'De volgende gebruikers zullen verwijderd worden uit de opgegeven groep:'
		$VBGroupsToBeDeleted = 'De gebruiker zal verwijderd worden uit de volgende groep(en):'
		
		$VBRemovingUser = 'Gebruiker verwijderen uit de groep'
		$VBRemovingUserSuccess = 'Gebruiker succesvol verwijderd uit de groep'
		$VBRemovingUserFailed = 'Fout bij het verwijderen van de gebruiker uit de groep'
		$VBRemovingUserCancelled = 'Verwijderen van gebruiker geannuleerd.'
		
		$VBGroupsCopied = 'Geselecteerde groep(en) gekopieerd naar klembord.'
		$VBGroupsExported = 'Geselecteerde groep(en) geëxporteerd.'
		
		$VBUsersCopied = 'Geselecteerde gebruiker(s) gekopieerd naar klembord.'
		
		$MsgBody = 'Weet je zeker dat u de groep(en) wilt verwijderen?'
		$MsgBodyExport = 'Wil je alle groepen exporteren?'
		
		##### Add Group to User Form
		$VBGroupToBeAdded = 'Groep toevoegen:'
		$VBGroupToBeAddedCancelled = 'Groep(en) toevoegen geannuleerd.'
		$VBGroupToBeAddedSuccess = 'Groep succesvol toegevoegd'
		$VBGroupToBeAddedFailed = 'Kan groep niet toevoegen'
		$VBGroupAlreadyAdded = 'Geselecteerde groep is al toegevoegd.'
		
		#### Groups Tab
		$VBGRPNoGroupName = 'Vul alsjeblieft een groepsnaam in waarop je wilt zoeken.'
		$VBGRPNoGroupFound = 'Geen groep gevonden met de volgende naam'
		$VBGRPNoMember = 'Geen lid/leden geselecteerd. Selecteer 1 of meerdere leden.'
		$VBGRPMembersToBeDeleted = 'De volgende leden zullen verwijderd worden uit de groep:'
		
		$VBGRPRemovingMember = 'Verwijderen van lid uit groep'
		$VBGRPRemovingMemberSuccess = 'Lid succesvol verwijderd uit groep'
		$VBGRPRemovingMemberFailed = 'Fout bij het verwijderen van lid uit groep'
		$VBGRPRemovingMemberCancelled = 'Verwijderen van lid/leden geannuleerd.'
		
		$VBGRPMembersCopied = 'Geselecteerde lid/leden gekopieerd naar klembord.'
		$VBGRPMembersExported = 'Geselecteerde lid/leden geëxporteerd.'
		
		##### Add User to Group Form
		$VBUserToBeAdded = 'Gebruiker toevoegen:'
		$VBUserToBeAddedCancelled = 'Toevoegen gebruiker(s) geannuleerd.'
		$VBUserToBeAddedSuccess = 'Gebruiker succesvol toegevoegd.'
		$VBUserToBeAddedFailed = 'Kan gebruiker niet toevoegen.'
		$VBUserAlreadyAdded = 'De geselecteerde gebruiker is al toegevoegd.'
		
		$GRPMsgBody = 'Weet je zeker dat je het lid/de leden wilt verwijderen?'
		$GRPMsgBodyExport = 'Wil je alle leden exporteren?'
		#endregion Dutch Language
	}
	elseif ($Culture -like "de-*")
	{
		#region German Language
		$Language = "German"
		$SSUpdateLanguageList = 'Sprachliste aktualisieren'
		$SSSettingFormControls = 'Steuerelemente initialisieren'
		$SSSettingUILanguage = 'Sprache festlegen'
		$SSConfiguringRemotePSSession = 'PS-Remoting-Sitzung konfigurieren'
		$SSRefreshingMembership = 'Gruppenmitgliedschaft aktualisieren...'
		$SSSearchingGroups = 'Gruppen suchen...'
		$SSSearchingUsers = 'Benutzers suchen...'
		
		### General Form Controls
		$FormLabel = "Active Directory Benutzer und Gruppen v$($VersionNumber) - $($OrgName)"
		$SSTitle = 'Wird geladen...'
		
		$ButtonSearch = "&Suche"
		$ButtonAdd = "&Hinzufügen"
		$ButtonRemove = "&Entfernen"
		$ButtonAgain = "&Nochmal"
		$ButtonCopy = "&Kopieren"
		$ButtonExport = "&Export"
		$ButtonExit = "E&xit"
		$LabelVerboseOutput = 'Erweiterte Ausgabe:'
		$LabelLanguage = 'Sprache:'
		$LabelDisplayName = 'Anzeigename:'
		$GroupSearchControls = 'Suchsteuerelemente'
		$TabPage0 = 'Benutzer'
		$TabPage1 = 'Gruppen'
		$MembersTabPage0 = 'Mitglieder'
		$MembersTabPage1 = 'Mitglied von'
		
		### Form Labels
		#### Users Tab
		$LabelUsername = 'Benutzername:'
		$LabelGroupMembership = 'Gruppenmitgliedschaften:'
		$LabelDepartment = 'Abteilung:'
		$LabelFunction = 'Funktion:'
		$LabelSamAccountName = 'Benutzername:'
		$LabelUserPrincipalName = 'Benutzerprinzipalname:'
		$LabelEmailAddress = 'E-Mail-Adresse:'
		$GroupUserDetails = 'Benutzerdetails'
		
		##### Add Group to User Form
		$AFFormName = 'Active Directory Benutzer - Gruppe(n) hinzufügen'
		$AFlblADGroup = 'AD-Gruppe:'
		$AFbtnAdd = "&Add"
		$AFbtnCancel = "A&bbrechen"
		$AFbtnRetrieve = "Ab&rufen"
		$AFNoADGroup = 'Kein Gruppenname eingegeben.'
		$AFADGroupFound = 'Gruppe(n) gefunden:'
		$AFADGroupInvalid = 'Keine Gruppe(n) gefunden:'
		$AFGroupsFound = 'Gruppe(n) gefunden:'
		$AFGroupsToBeAdded = 'Gruppe(n) hinzufügen:'
		
		#### Groups Tab
		$GRPLabelGroupname = 'Gruppenname:'
		$GRPLabelGroupCategory = 'Gruppenkategorie:'
		$GRPLabelGroupScope = 'Gruppenbereich:'
		$GRPLabelGroupInfo = 'Gruppeninfo:'
		$GRPLabelManagedBy = 'Verwaltet von:'
		$GRPLabelDirectMemberCount = 'Anzahl der direkten Mitglieder:'
		$GRPLabelNestedGroupsCount = 'Anzahl der verschachtelten Gruppen:'
		$GRPLabelMemberOfCount = 'Anzahl Mitglieder von:'
		$GRPGroupGroupDetails = 'Gruppe details'
		
		#### Select Group Form
		$SGFormName = 'Active Directory Gruppen - Gruppe wählen'
		$SGFormNameUsers = 'Active Directory Benutzer - Benutzer wählen'
		$SGbtnCancel = "&Stornieren"
		$SGbtnOK = "&OK"
		
		#### Add User to Group Form
		$AUFormName = 'Active Directory-Gruppen - Benutzer hinzufügen'
		$AUlblADUser = 'AD-Benutzer:'
		$AUNoADUser = 'Kein Benutzername eingegeben.'
		$AUUserFound = 'Benutzer gefunden:'
		$AUADUserInvalid = 'Keine Benutzer gefunden.'
		$AUUsersFound = 'Benutzer gefunden:'
		$AUUsersToBeAdded = 'Benutzer hinzufügen:'
		
		### ToolTip Text
		#### General Form Controls
		$TTcmbLanguage = 'Hier können Sie die Formularsprache auswählen.'
		$TTbtnExit = 'Diese Aktion schließt das Formular.'
		$TTrtbOutput = 'Dies ist der Protokollbildschirm.'
		
		#### Users Tab
		$TTtxtUsername = 'Geben Sie hier den Benutzernamen des Benutzers ein, den Sie suchen möchten.'
		$TTbtnSearch = "Diese Aktion sollte zuerst ausgeführt werden. Wenn der Benutzer gefunden wird, sind die Schaltflächen `'Hinzufügen`', `'Entfernen`', `'Kopieren`' und `'Exportieren`' verfügbar."
		$TTbtnAdd = 'Diese Aktion öffnet ein Popup, in dem eine Active Directory-Gruppe gefunden wird, zu der der Benutzer hinzugefügt wird.'
		$TTbtnRemove = 'Diese Aktion entfernt den Benutzer aus den ausgewählten Gruppen.'
		$TTbtnAgain = 'Diese Aktion bringt das Formular in den Ausgangszustand zurück.'
		$TTbtnCopy = 'Diese Aktion kopiert die ausgewählten Gruppen in die Zwischenablage.'
		$TTbtnExport = 'Diese Aktion exportiert die ausgewählten Gruppen in eine Datei.'
		$TTlstADGroups = 'In diesem Bildschirm werden die aktuellen Gruppen angezeigt, zu denen der angegebene Benutzer gehört.'
		
		##### Add AD Group Form
		$TTAddtxtADGroup = 'Bitte füllen Sie dieses Feld mit der AD-Gruppe (oder dem Platzhalter) aus, nach der Sie suchen möchten.'
		$TTAddlstAddADGroups = "Diese Dropdown-Liste enthält die gefundenen Gruppen basierend auf den Suchkriterien.`r`nDurch Doppelklicken auf eine Gruppe wird sie an die andere Dropdown-Liste gesendet, die dem Benutzer hinzugefügt werden kann.`r`nCtrl + A wählt alle gefundenen Gruppen aus.`r`nDurch Drücken von Ctrl+C werden die ausgewählten Gruppen in die Zwischenablage kopiert."
		$TTAddlstGroupsToBeAdded = "Dieses Listenfeld enthält die Gruppen, die dem angegebenen Benutzer hinzugefügt werden können.`r`nDurch Doppelklicken auf eine Gruppe wird sie aus der Liste entfernt.`r`nDurch Drücken von Ctrl+V im Listenfeld ausgewählt ist, wird in die Zwischenablage Inhalt eingefügt."
		$TTAddbtnAdd = 'Diese Aktion durchläuft die ausgewählten Gruppen und fügt sie dem angegebenen Benutzer hinzu.'
		$TTAddbtnCancel = 'Diese Aktion bricht das Hinzufügen ab und schließt das Formular.'
		$TTAddbtnAgain = 'Diese Aktion bringt das Formular in seinen ursprünglichen Zustand zurück.'
		$TTAddbtnRetrieve = 'Diese Aktion sucht basierend auf Ihrer Eingabe nach Active Directory-Gruppen.'
		$TTAddbtnLeft = "Diese Aktion entfernt die ausgewählten Gruppen aus der Liste `'Gruppe(n) hinzufügen`'."
		$TTAddbtnRight = "Mit dieser Aktion werden die ausgewählten Gruppen aus der Liste `'Gefundene Gruppe(n)`' zur Liste `'Gruppe(n) hinzufügen`" hinzugefügt."
		$TTAddrtbAddOutput = 'Dies ist der Protokollbildschirm.'
		
		#### Groups Tab
		$TTGRPtxtGroupname = 'Geben Sie hier den Gruppennamen der Gruppe ein, nach der Sie suchen möchten.'
		$TTGRPbtnSearch = "Diese Aktion sollte zuerst ausgeführt werden. Wenn die Gruppe gefunden wurde, sind die Schaltflächen `'Hinzufügen`', `'Löschen`', `'Kopieren`' und `'Exportieren`' verfügbar."
		$TTGRPbtnAdd = 'Diese Aktion generiert ein Popup, in dem ein Active Directory-Benutzer gefunden wird, der der Gruppe hinzugefügt wird.'
		$TTGRPbtnRemove = 'Diese Aktion entfernt die ausgewählten Benutzer aus der Gruppe.'
		$TTGRPbtnAgain = 'Diese Aktion bringt das Formular in den Ausgangszustand zurück.'
		$TTGRPbtnCopy = 'Diese Aktion kopiert die ausgewählten Benutzer in die Zwischenablage.'
		$TTGRPbtnExport = 'Diese Aktion exportiert die ausgewählten Benutzer in eine Datei.'
		$TTlstGroupMembers = 'In diesem Bildschirm werden die aktuellen Mitglieder der angegebenen Gruppe angezeigt.'
		$TTlstGroupMemberOf = 'In diesem Bildschirm werden die aktuellen Gruppen angezeigt, zu denen die angegebene Gruppe gehört.'
		
		##### Select Group Form
		$TTSGbtnCancel = 'Diese Aktion bricht die Gruppenauswahl ab.'
		$TTSGbtnOK = 'Diese Aktion verwendet die markierte Gruppe für weitere Aktionen.'
		$TTSGlstSelectObject = "Diese Liste enthält die Gruppen, die basierend auf der Eingabe gefunden wurden.`r`nBitte wählen Sie eine Gruppe aus, um fortzufahren."
		
		$TTAddUtxtADUser = 'Bitte füllen Sie dieses Feld mit dem AD-Benutzer (oder Platzhalter) aus, nach dem Sie suchen möchten.'
		$TTAddUlstAddADUsers = "Diese Dropdown-Liste enthält die gefundenen Benutzer basierend auf den Suchkriterien.`r`n Durch Doppelklicken auf einen Benutzer wird diese an die andere Dropdown-Liste gesendet, die dem Benutzer hinzugefügt werden kann.`r`nCtrl+A wählt alle gefundenen Benutzer aus.`r`nDurch Drücken von Ctrl+C werden die ausgewählten Benutzer in die Zwischenablage kopiert."
		$TTAddUlstUsersToBeAdded = "Dieses Listenfeld enthält die Benutzer, die bereit sind, der angegebenen Gruppe hinzugefügt zu werden.`r`nDurch Doppelklicken auf einen Benutzer wird dieser aus der Liste entfernt.`r`nDurch Drücken von Ctrl+V im Listenfeld ausgewählt ist, wird der Inhalt der Zwischenablage eingefügt."
		$TTAddUbtnAdd = 'Diese Aktion führt die ausgewählten Benutzer durch und fügt sie der angegebenen Gruppe hinzu.'
		$TTAddUbtnCancel = 'Diese Aktion bricht das Hinzufügen ab und schließt das Formular.'
		$TTAddUbtnAgain = 'Diese Aktion bringt das Formular in seinen ursprünglichen Zustand zurück.'
		$TTAddUbtnRetrieve = 'Diese Aktion sucht anhand Ihrer Eingabe nach Active Directory-Benutzern.'
		$TTAddUbtnLeft = "Diese Aktion entfernt die ausgewählten Benutzer aus der Liste `'Benutzer hinzufügen`'."
		$TTAddUbtnRight = "Diese Aktion fügt die ausgewählten Benutzer aus der Liste `'Gefundene Benutzer`' zur Liste der hinzuzufügenden Benutzer hinzu."
		$TTAddUrtbAddOutput = 'Dies ist der Protokollbildschirm.'
		
		### Verbose Output
		#### General Strings
		$VBInfo = 'INFO:'
		$VBWarning = 'WARNUNG:'
		$VBError = 'FEHLER:'
		$VBSuccess = 'ERFOLG:'
		$VBModuleCheck = 'Prüfen'
		$VBModule = 'Modul'
		$VBModuleAlreadyImported = 'Modul wurde bereits importiert.'
		$VBModuleAvailable = 'Modul verfügbar. Importieren...'
		$VBModuleImported = 'Modul importiert.'
		$VBModuleSearch = 'Suche nach'
		$VBModuleOnline = 'Modul in der Online-Galerie.'
		$VBModuleFound = 'Modul gefunden. Installieren...'
		$VBModuleInstalled = 'Modul installiert. Importieren...'
		$VBModuleFailed = 'Modul nicht importiert, nicht verfügbar und nicht in der Online-Galerie gefunden.'
		$VBPSSessionSetup = 'PS-Remoting-Sitzung einrichten.'
		$VBPSSessionSetupSuccess = 'PS-Remoting-Verbindung hergestellt.'
		$VBPSSessionSetupFailed = 'PS-Remoting konnte keine Verbindung herstellen. Bitte kontaktieren Sie Ihren Systemadministrator.'
		$VBPSSessionImport = 'PS-Remoting-Sitzung importieren.'
		$VBPSSessionImportSuccess = 'Importierte PS-Remoting-Sitzung.'
		$VBPSSessionImportFailed = 'PS-Remoting konnte Sitzung nicht importieren. Bitte kontaktieren Sie Ihren Systemadministrator.'
		$VBPSSessionInvalid = 'PS Remoting Keine gültige Sitzung gefunden. Bitte kontaktieren Sie Ihren Systemadministrator.'
		$VBButtonNotImplemented = 'Diese Schaltfläche wurde noch nicht implementiert.'
		$VBNoCBContent = 'Die Zwischenablage ist leer.'
		$VBCBPasted = 'Inhalt der Zwischenablage eingefügt.'
		$VBTooFewChar = "Es wurden zu wenige Zeichen eingegeben:`r`n`tBenutzernamen: mindestens 2`r`n`tGruppennamen: mindestens 4`r`n"
		
		#### Users Tab
		$VBNoUserName = 'Bitte geben Sie einen Benutzernamen ein, nach dem Sie suchen möchten.'
		$VBNoUserFound = 'Kein Benutzer mit folgendem Benutzernamen gefunden'
		$VBNoGroup = 'Keine Gruppe(n) ausgewählt. Wählen Sie eine oder mehrere Gruppen aus.'
		$VBUsersToBeDeleted = 'Die folgenden Benutzer werden aus der angegebenen Gruppe entfernt:'
		$VBGroupsToBeDeleted = 'Der Benutzer wird aus den folgenden Gruppe(n) entfernt:'
		
		$VBRemovingUser = 'Benutzer aus Gruppe entfernen'
		$VBRemovingUserSuccess = 'Benutzer erfolgreich aus der Gruppe entfernt'
		$VBRemovingUserFailed = 'Fehler beim Entfernen des Benutzers aus der Gruppe'
		$VBRemovingUserCancelled = 'Löschen des Benutzers abgebrochen.'
		
		$VBGroupsCopied = 'Ausgewählte Gruppe(n) in die Zwischenablage kopiert.'
		$VBGroupsExported = 'Ausgewählte Gruppe(n) exportiert.'
		
		$VBUsersCopied = 'Ausgewählte Benutzer in die Zwischenablage kopiert.'
		
		$MsgBody = 'Möchten Sie die Gruppe(n) wirklich löschen?'
		$MsgBodyExport = 'Möchten Sie alle Gruppen exportieren?'
		
		##### Add Group to User Form
		$VBGroupToBeAdded = 'Gruppen hinzufügen:'
		$VBGroupToBeAddedCancelled = 'Gruppe(n) hinzufügen abgebrochen.'
		$VBGroupToBeAddedSuccess = 'Gruppe erfolgreich hinzugefügt'
		$VBGroupToBeAddedFailed = 'Gruppe kann nicht hinzugefügt werden'
		$VBGroupAlreadyAdded = 'Ausgewählte Gruppe wurde bereits hinzugefügt.'
		
		#### Groups Tab
		$VBGRPNoGroupName = 'Geben Sie einen Gruppennamen ein, nach dem Sie suchen möchten.'
		$VBGRPNoGroupFound = 'Keine Gruppe mit folgendem Namen gefunden'
		$VBGRPNoMember = 'Keine Mitglieder ausgewählt. Wählen Sie ein oder mehrere Mitglieder aus. '
		$VBGRPMembersToBeDeleted = 'Die folgenden Mitglieder werden aus der Gruppe entfernt:'
		
		$VBGRPRemovingMember = 'Mitglied aus Gruppe entfernen'
		$VBGRPRemovingMemberSuccess = 'Mitglied erfolgreich aus Gruppe entfernt'
		$VBGRPRemovingMemberFailed = 'Fehler beim Entfernen des Mitglieds aus der Gruppe'
		$VBGRPRemovingMemberCancelled = 'Entfernen des Mitglieds abgebrochen.'
		
		$VBGRPMembersCopied = 'Ausgewählte Mitglieder in die Zwischenablage kopiert.'
		$VBGRPMembersExported = 'Ausgewählte Mitglieder exportiert.'
		
		###### Add User to Group Form
		$VBUserToBeAdded = 'Benutzer hinzufügen:'
		$VBUserToBeAddedCancelled = 'Benutzer hinzugefügen abgebrochen.'
		$VBUserToBeAddedSuccess = 'Benutzer erfolgreich hinzugefügt.'
		$VBUserToBeAddedFailed = 'Benutzer kann nicht hinzugefügt werden.'
		$VBUserAlreadyAdded = 'Ausgewählter Benutzer wurde bereits hinzugefügt.'
		
		$GRPMsgBody = 'Möchten Sie die Mitglieder wirklich löschen?'
		$GRPMsgBodyExport = 'Möchten Sie alle Mitglieder exportieren?'
		#endregion German Language
	}
	else
	{
		#region English language
		$Language = "English"
		
		### General Form Controls
		$FormLabel = "Active Directory Users and Groups v$($VersionNumber) - $($OrgName)"
		$SSTitle = 'Loading...'
		$SSUpdateLanguageList = 'updating language list'
		$SSSettingFormControls = 'initializing form controls'
		$SSSettingUILanguage = 'setting language'
		$SSConfiguringRemotePSSession = 'configuring PS-Remoting session'
		$SSRefreshingMembership = 'Refreshing group membership'
		$SSSearchingGroups = 'Searching groups...'
		$SSSearchingUsers = 'Searching users...'
		
		$ButtonSearch = "&Search"
		$ButtonAdd = "&Add"
		$ButtonRemove = "&Remove"
		$ButtonAgain = "&Again"
		$ButtonCopy = "&Copy"
		$ButtonExport = "&Export"
		$ButtonExit = "E&xit"
		$LabelVerboseOutput = 'Verbose output:'
		$LabelLanguage = 'Language:'
		$LabelDisplayName = 'Displayname:'
		$GroupSearchControls = 'Search Controls'
		$TabPage0 = 'Users'
		$TabPage1 = 'Groups'
		$MembersTabPage0 = 'Members'
		$MembersTabPage1 = 'Member Of'
		
		### Form Labels
		#### Users Tab
		$LabelUsername = 'Username:'
		$LabelGroupMembership = 'Group memberships:'
		$LabelDepartment = 'Department:'
		$LabelFunction = 'Function:'
		$LabelSamAccountName = 'SamAccountName:'
		$LabelUserPrincipalName = 'User Principal Name:'
		$LabelEmailAddress = 'Email address:'
		$GroupUserDetails = 'User details'
		
		#### Add Group to User Form
		$AFFormName = 'Active Directory Users - Add Group(s)'
		$AFlblADGroup = 'AD group:'
		$AFbtnAdd = "&Add"
		$AFbtnCancel = "&Cancel"
		$AFbtnRetrieve = "&Retrieve"
		$AFNoADGroup = 'No group name entered.'
		$AFADGroupFound = 'Found group(s):'
		$AFADGroupInvalid = 'No group(s) found.'
		$AFGroupsFound = 'Group(s) found:'
		$AFGroupsToBeAdded = 'Group(s) to be added:'
		
		#### Groups Tab
		$GRPLabelGroupname = 'Groupname:'
		$GRPLabelGroupCategory = 'Group Category:'
		$GRPLabelGroupScope = 'Group Scope:'
		$GRPLabelGroupInfo = 'Group Info:'
		$GRPLabelManagedBy = 'Managed by:'
		$GRPLabelDirectMemberCount = 'Number of Direct members:'
		$GRPLabelNestedGroupsCount = 'Number of Nested groups:'
		$GRPLabelMemberOfCount = 'MemberOf Count:'
		$GRPGroupGroupDetails = 'Group details'
		
		#### Select Group Form
		$SGFormName = 'Active Directory Groups - Select Group'
		$SGFormNameUsers = 'Active Directory Users - Select User'
		$SGbtnCancel = "&Cancel"
		$SGbtnOK = "&OK"
		
		#### Add User to Group Form
		$AUFormName = 'Active Directory Groups - Add user(s)'
		$AUlblADUser = 'AD User:'
		$AUNoADUser = 'No username entered.'
		$AUUserFound = 'User(s) found:'
		$AUADUserInvalid = 'No user(s) found.'
		$AUUsersFound = 'Found user(s):'
		$AUUsersToBeAdded = 'Add User(s):'
		
		### Tooltip text
		#### General Form Controls
		$TTcmbLanguage ='Here you can choose the form language.'
		$TTbtnExit = 'This action will close the form.'
		$TTrtbOutput = 'This is the log screen.'
		
		#### Users Tab
		$TTtxtUsername = 'Enter the username of the user you want to look up here.'
		$TTbtnSearch = "This action should be performed first. When the user is found, the `'Add`', `'Remove`', `'Copy`' and `'Export`' buttons will become available."
		$TTbtnAdd = 'This action will generate a popup where an Active Directory group can be found to which the user will be added.'
		$TTbtnRemove = 'This action will remove the user from the selected group(s).'
		$TTbtnAgain = 'This action will return the form to its initial state.'
		$TTbtnCopy = 'This action copies the selected group(s) to the clipboard.'
		$TTbtnExport = 'This action exports the selected group(s) to a file.'
		$TTlstADGroups = 'This screen will display the current groups that the specified user is a member of.'
		
		##### Add AD Group Form
		$TTAddtxtADGroup = 'Please fill in this field with the AD group (or wildcard) you want to search for.'
		$TTAddlstAddADGroups = "This drop-down list contains the groups found, based on the search criteria.`r`nDouble clicking on a group will send it to the other drop-down list, ready to be added to the user.`r`nCtrl+A selects all found groups.`r`nPressing Ctrl+C copies the selected groups to the clipboard."
		$TTAddlstGroupsToBeAdded = "This list box contains the groups that are ready to be added to the specified user.`r`nDouble clicking on a group will remove it from the list.`r`nBy pressing Ctrl+V while the list box is selected, the clipboard content is pasted."
		$TTAddbtnAdd = 'This action will loop through the selected groups and add them to the specified user.'
		$TTAddbtnCancel = 'This action will cancel the add and close the form.'
		$TTAddbtnAgain = 'This action will return the form to its original state.'
		$TTAddbtnRetrieve = 'This action searches for Active Directory groups based on your input.'
		$TTAddbtnLeft = "This action will remove the selected group(s) from the `'Add Group(s)`' list."
		$TTAddbtnRight = "This action will add the selected group(s) from the `'Found Group(s)`' list to the `'Add Group(s)`' list."
		$TTAddrtbAddOutput = 'This is the log screen.'
		
		#### Groups Tab
		$TTGRPtxtGroupname = 'Enter the group name of the group you want to look up here.'
		$TTGRPbtnSearch = "This action should be performed first. When the group is found, the `'Add`', `'Delete`', `'Copy`' and `'Export`' buttons will become available."
		$TTGRPbtnAdd = 'This action will generate a popup where an Active Directory user can be found to be added to the group.'
		$TTGRPbtnRemove = 'This action will remove the selected user(s) from the group.'
		$TTGRPbtnAgain = 'This action will return the form to its initial state.'
		$TTGRPbtnCopy = 'This action copies the selected user(s) to the clipboard.'
		$TTGRPbtnExport = 'This action will export the selected user(s) to a file.'
		$TTlstGroupMembers = 'This screen will display the current members of the specified group.'
		$TTlstGroupMemberOf = 'This screen will display the current groups of which the specified group is a member.'
		
		##### Select Group Form
		$TTSGbtnCancel = 'This action cancels the group selection.'
		$TTSGbtnOK = 'This action will use the highlighted group for further actions.'
		$TTSGlstSelectObject = "This list contains the groups found based on input.`r`nPlease select a group to proceed."
		
		$TTAddUtxtADUser = 'Please fill in this field with the AD user (or wildcard) you want to search for.'
		$TTAddUlstAddADUsers = "This drop-down list contains the users found, based on the search criteria.`r`nDouble-clicking a user will send it to the other drop-down list, ready to be added to the user.`r`nCtrl+A selects all found users.`r`n Pressing Ctrl+C copies the selected users to the clipboard."
		$TTAddUlstUsersToBeAdded = "This list box contains the users who are ready to be added to the specified group.`r`nDouble clicking on a user will remove it from the list.`r`nBy pressing Ctrl+V while the list box is selected, the contents of the clipboard are pasted."
		$TTAddUbtnAdd = 'This action will run through the selected users and add them to the specified group.'
		$TTAddUbtnCancel = 'This action will cancel the add and close the form.'
		$TTAddUbtnAgain = 'This action will return the form to its original state.'
		$TTAddUbtnRetrieve = 'This action searches for Active Directory users based on your input.'
		$TTAddUbtnLeft = "This action will remove the selected user(s) from the `'Add User(s)`' list."
		$TTAddUbtnRight = "This action will add the selected user(s) from the `'Found user(s)`' list to the `'User(s) to be added`' list."
		$TTAddUrtbAddOutput = 'This is the log screen.'
		
		### Verbose output
		#### General strings
		$VBInfo = 'INFO:'
		$VBWarning = 'WARNING:'
		$VBError = 'ERROR:'
		$VBSuccess = 'SUCCESS:'
		$VBModuleCheck = 'Check for'
		$VBModule = 'module'
		$VBModuleAlreadyImported = 'Module has already been imported.'
		$VBModuleAvailable = 'module available. Import...'
		$VBModuleImported = 'module imported.'
		$VBModuleSearch = 'Search for'
		$VBModuleOnline = 'module in online gallery.'
		$VBModuleFound = 'module found. To install...'
		$VBModuleInstalled = 'module installed. Import...'
		$VBModuleFailed = 'Module not imported, not available, and not found in the online gallery.'
		$VBPSSessionSetup = 'Setting up PS remoting session.'
		$VBPSSessionSetupSuccess = 'PS remoting connection established.'
		$VBPSSessionSetupFailed = 'PS remoting failed to connect. Please contact your System Administrator '
		$VBPSSessionImport = 'Import PS remoting session.'
		$VBPSSessionImportSuccess = 'Imported PS remoting session.'
		$VBPSSessionImportFailed = 'PS remoting failed to import session. Please contact your System Administrator. '
		$VBPSSessionInvalid = 'PS Remoting No valid session found. Please contact your System Administrator '
		$VBButtonNotImplemented = 'This button has not been implemented yet.'
		$VBNoCBContent = 'The clipboard is empty.'
		$VBCBPasted = 'Clipboard content pasted.'
		$VBTooFewChar = "Too few characters entered.`r`n`tUsername: at least 2`r`n`tGroupname: at least 4`r`n"
		
		#### Users Tab
		$VBNoUserName = 'Please enter a username you want to search for.'
		$VBNoUserFound = 'No user found with the following username'
		$VBNoGroup = 'No group(s) selected. Select 1 or more groups.'
		$VBUsersToBeDeleted = 'The following users will be removed from the specified group:'
		$VBGroupsToBeDeleted = 'The user will be removed from the following group(s):'

		$VBRemovingUser = 'Remove user from group'
		$VBRemovingUserSuccess = 'User successfully removed from the group'
		$VBRemovingUserFailed = 'Error removing user from group'
		$VBRemovingUserCancelled = 'User deletion canceled.'

		$VBGroupsCopied = 'Selected group(s) copied to clipboard.'
		$VBGroupsExported = 'Selected group(s) exported.'
		
		$VBUsersCopied = 'Selected user(s) copied to clipboard.'
		
		$MsgBody = 'Are you sure you want to delete the group(s)?'
		$MsgBodyExport = 'Do you want to export all groups?'

		##### Add Group to User Form
		$VBGroupToBeAdded = 'Add Group:'
		$VBGroupToBeAddedCancelled = 'Add group(s) canceled.'
		$VBGroupToBeAddedSuccess = 'Group added successfully'
		$VBGroupToBeAddedFailed = 'Unable to add group'
		$VBGroupAlreadyAdded = 'Selected group has already been added.'

		#### Groups Tab
		$VBGRPNoGroupName = 'Please enter a group name you want to search for.'
		$VBGRPNoGroupFound = 'No group found with the following name'
		$VBGRPNoMember = 'No member(s) selected. Select 1 or more members. '
		$VBGRPMembersToBeDeleted = 'The following member(s) will be removed from the group:'

		$VBGRPRemovingMember = 'Remove member from group'
		$VBGRPRemovingMemberSuccess = 'Member successfully removed from group'
		$VBGRPRemovingMemberFailed = 'Error removing member from group'
		$VBGRPRemovingMemberCancelled = 'Removing member canceled.'

		$VBGRPMembersCopied = 'Selected member(s) copied to clipboard.'
		$VBGRPMembersExported = 'Selected member(s) exported.'
		
		##### Add User to Group Form
		$VBUserToBeAdded = 'Add User:'
		$VBUserToBeAddedCancelled = 'Add user(s) canceled.'
		$VBUserToBeAddedSuccess = 'User added successfully.'
		$VBUserToBeAddedFailed = 'Unable to add user.'
		$VBUserAlreadyAdded = 'Selected user has already been added.'

		$GRPMsgBody = 'Are you sure you want to delete the member(s)?'
		$GRPMsgBodyExport = 'Do you want to export all members?'
		#endregion English Language
	}
	
	#region Export Variable Strings
	$LanguageObject = @()
	
	$Properties = @{
		Language					 = $Language
		FormLabel				     = $FormLabel
		SSTitle						 = $SSTitle
		SSUpdateLanguageList		 = $SSUpdateLanguageList
		SSSettingFormControls		 = $SSSettingFormControls
		SSSettingUILanguage			 = $SSSettingUILanguage
		SSConfiguringRemotePSSession = $SSConfiguringRemotePSSession
		SSRefreshingMembership		 = $SSRefreshingMembership
		SSSearchingGroups		     = $SSSearchingGroups
		SSSearchingUsers		     = $SSSearchingUsers
		ButtonSearch				 = $ButtonSearch
		ButtonAdd				     = $ButtonAdd
		ButtonRemove				 = $ButtonRemove
		ButtonAgain				     = $ButtonAgain
		ButtonCopy				     = $ButtonCopy
		ButtonExport				 = $ButtonExport
		ButtonExit				     = $ButtonExit
		LabelVerboseOutput		     = $LabelVerboseOutput
		LabelLanguage			     = $LabelLanguage
		LabelDisplayName			 = $LabelDisplayName
		GroupSearchControls		     = $GroupSearchControls
		TabPage0					 = $TabPage0
		TabPage1					 = $TabPage1
		MembersTabPage0				 = $MembersTabPage0
		MembersTabPage1				 = $MembersTabPage1
		LabelUsername			     = $LabelUsername
		LabelGroupMembership		 = $LabelGroupMembership
		LabelDepartment			     = $LabelDepartment
		LabelFunction			     = $LabelFunction
		LabelSamAccountName		     = $LabelSamAccountName
		LabelUserPrincipalName	     = $LabelUserPrincipalName
		LabelEmailAddress		     = $LabelEmailAddress
		GroupUserDetails			 = $GroupUserDetails
		GRPLabelGroupname		     = $GRPLabelGroupname
		GRPLabelGroupCategory	     = $GRPLabelGroupCategory
		GRPLabelGroupScope		     = $GRPLabelGroupScope
		GRPLabelGroupInfo		     = $GRPLabelGroupInfo
		GRPLabelManagedBy		     = $GRPLabelManagedBy
		GRPLabelDirectMemberCount    = $GRPLabelDirectMemberCount
		GRPLabelNestedGroupsCount    = $GRPLabelNestedGroupsCount
		GRPLabelMemberOfCount	  	 = $GRPLabelMemberOfCount
		GRPGroupGroupDetails		 = $GRPGroupGroupDetails
		SGFormName					 = $SGFormName
		SGFormNameUsers			 	 = $SGFormNameUsers
		SGbtnOK						 = $SGbtnOK
		SGbtnCancel					 = $SGbtnCancel
		AUFormName				     = $AUFormName
		AUlblADUser				 	 = $AUlblADUser
		AUNoADUser				     = $AUNoADUser
		AUUserFound				 	 = $AUUserFound
		AUADUserInvalid			 	 = $AUADUserInvalid
		AUUsersFound			     = $AUUsersFound
		AUUsersToBeAdded		     = $AUUsersToBeAdded
		TTcmbLanguage			     = $TTcmbLanguage
		TTbtnExit				     = $TTbtnExit
		TTrtbOutput				     = $TTrtbOutput
		TTtxtUsername			     = $TTtxtUsername
		TTbtnSearch				     = $TTbtnSearch
		TTbtnAdd					 = $TTbtnAdd
		TTbtnRemove				     = $TTbtnRemove
		TTbtnAgain				     = $TTbtnAgain
		TTbtnCopy				     = $TTbtnCopy
		TTbtnExport				     = $TTbtnExport
		TTlstADGroups			     = $TTlstADGroups
		TTAddtxtADGroup			     = $TTAddtxtADGroup
		TTAddlstAddADGroups		     = $TTAddlstAddADGroups
		TTAddlstGroupsToBeAdded	     = $TTAddlstGroupsToBeAdded
		TTAddbtnAdd				     = $TTAddbtnAdd
		TTAddbtnCancel			     = $TTAddbtnCancel
		TTAddbtnAgain			     = $TTAddbtnAgain
		TTAddbtnRetrieve			 = $TTAddbtnRetrieve
		TTAddbtnLeft				 = $TTAddbtnLeft
		TTAddbtnRight			     = $TTAddbtnRight
		TTAddrtbAddOutput		     = $TTAddrtbAddOutput
		TTGRPtxtGroupname		     = $TTGRPtxtGroupname
		TTGRPbtnSearch			     = $TTGRPbtnSearch
		TTGRPbtnAdd				     = $TTGRPbtnAdd
		TTGRPbtnRemove			     = $TTGRPbtnRemove
		TTGRPbtnAgain			     = $TTGRPbtnAgain
		TTGRPbtnCopy				 = $TTGRPbtnCopy
		TTGRPbtnExport			     = $TTGRPbtnExport
		TTlstGroupMembers		     = $TTlstGroupMembers
		TTSGbtnOK					 = $TTSGbtnOK
		TTSGbtnCancel				 = $TTSGbtnCancel
		TTSGlstSelectObject			 = $TTSGlstSelectObject
		TTAddUtxtADUser			 	 = $TTAddUtxtADUser
		TTAddUlstAddADUsers		 	 = $TTAddUlstAddADUsers
		TTAddUlstUsersToBeAdded	 	 = $TTAddUlstUsersToBeAdded
		TTAddUbtnAdd			     = $TTAddUbtnAdd
		TTAddUbtnCancel				 = $TTAddUbtnCancel
		TTAddUbtnAgain			     = $TTAddUbtnAgain
		TTAddUbtnRetrieve		     = $TTAddUbtnRetrieve
		TTAddUbtnLeft			     = $TTAddUbtnLeft
		TTAddUbtnRight			     = $TTAddUbtnRight
		TTAddUrtbAddOutput		     = $TTAddUrtbAddOutput
		VBInfo					     = $VBInfo
		VBWarning				     = $VBWarning
		VBError					     = $VBError
		VBSuccess				     = $VBSuccess
		VBModuleCheck			     = $VBModuleCheck
		VBModule					 = $VBModule
		VBModuleAlreadyImported	     = $VBModuleAlreadyImported
		VBModuleAvailable		     = $VBModuleAvailable
		VBModuleImported			 = $VBModuleImported
		VBModuleSearch			     = $VBModuleSearch
		VBModuleOnline			     = $VBModuleOnline
		VBModuleFound			     = $VBModuleFound
		VBModuleInstalled		     = $VBModuleInstalled
		VBModuleFailed			     = $VBModuleFailed
		VBPSSessionSetup			 = $VBPSSessionSetup
		VBPSSessionSetupSuccess	     = $VBPSSessionSetupSuccess
		VBPSSessionSetupFailed	     = $VBPSSessionSetupFailed
		VBPSSessionImport		     = $VBPSSessionImport
		VBPSSessionImportSuccess	 = $VBPSSessionImportSuccess
		VBPSSessionImportFailed	     = $VBPSSessionImportFailed
		VBPSSessionInvalid		     = $VBPSSessionInvalid
		VBButtonNotImplemented	     = $VBButtonNotImplemented
		VBNoCBContent			     = $VBNoCBContent
		VBCBPasted				     = $VBCBPasted
		VBTooFewChar				 = $VBTooFewChar
		VBNoUserName				 = $VBNoUserName
		VBNoUserFound			     = $VBNoUserFound
		VBNoGroup				     = $VBNoGroup
		VBUsersToBeDeleted			 = $VBUsersToBeDeleted
		VBGroupsToBeDeleted		     = $VBGroupsToBeDeleted
		VBRemovingUser			     = $VBRemovingUser
		VBRemovingUserSuccess	     = $VBRemovingUserSuccess
		VBRemovingUserFailed		 = $VBRemovingUserFailed
		VBRemovingUserCancelled	     = $VBRemovingUserCancelled
		VBGroupsCopied			     = $VBGroupsCopied
		VBGroupsExported			 = $VBGroupsExported
		VBUsersCopied				 = $VBUsersCopied
		VBUserToBeAdded			 	 = $VBUserToBeAdded
		VBUserToBeAddedCancelled     = $VBUserToBeAddedCancelled
		VBUserToBeAddedSuccess	     = $VBUserToBeAddedSuccess
		VBUserToBeAddedFailed	     = $VBUserToBeAddedFailed
		VBUserAlreadyAdded		     = $VBUserAlreadyAdded
		MsgBody					     = $MsgBody
		MsgBodyExport			     = $MsgBodyExport
		AFFormName				     = $AFFormName
		AFlblADGroup				 = $AFlblADGroup
		AFbtnAdd					 = $AFbtnAdd
		AFbtnCancel				     = $AFbtnCancel
		AFbtnRetrieve			     = $AFbtnRetrieve
		AFNoADGroup				     = $AFNoADGroup
		AFADGroupFound			     = $AFADGroupFound
		AFADGroupInvalid			 = $AFADGroupInvalid
		AFGroupsFound			     = $AFGroupsFound
		AFGroupsToBeAdded		     = $AFGroupsToBeAdded
		VBGroupToBeAdded			 = $VBGroupToBeAdded
		VBGroupToBeAddedCancelled    = $VBGroupToBeAddedCancelled
		VBGroupToBeAddedSuccess	     = $VBGroupToBeAddedSuccess
		VBGroupToBeAddedFailed	     = $VBGroupToBeAddedFailed
		VBGroupAlreadyAdded		     = $VBGroupAlreadyAdded
		VBGRPNoGroupName			 = $VBGRPNoGroupName
		VBGRPNoGroupFound		     = $VBGRPNoGroupFound
		VBGRPNoMember			     = $VBGRPNoMember
		VBGRPMembersToBeDeleted	     = $VBGRPMembersToBeDeleted
		VBGRPRemovingMember		     = $VBGRPRemovingMember
		VBGRPRemovingMemberSuccess   = $VBGRPRemovingMemberSuccess
		VBGRPRemovingMemberFailed    = $VBGRPRemovingMemberFailed
		VBGRPRemovingMemberCancelled = $VBGRPRemovingMemberCancelled
		VBGRPMembersCopied		     = $VBGRPMembersCopied
		VBGRPMembersExported		 = $VBGRPMembersExported
		GRPMsgBody				     = $GRPMsgBody
		GRPMsgBodyExport			 = $GRPMsgBodyExport
	}
	
	$LanguageObject = New-Object -TypeName PSObject -Property $Properties | Select-Object *
	
	return $LanguageObject
	#endregion Export Variable Strings
}

function Get-ADUserObject
{
	Param (
		[Parameter(Mandatory = $true)]
		[string]$Search,
		[Parameter(Mandatory = $true)]
		[string]$FilterType
	)
	
	if ($RemoteADModule)
	{
		$UserCount = Invoke-Command -Session $Session -ScriptBlock {
			Param ($search,
				$filter)
			Import-Module ActiveDirectory
			Get-ADUser -Filter { $filter -like $search } -Properties * -ErrorAction Stop | Sort-Object SamAccountName
		} -ArgumentList $Search, $FilterType
		#Write-Host "Remote PowerShell"
	}
	else
	{
		$UserCount = Get-ADUser -Filter { $FilterType -like $Search } -Properties * -ErrorAction Stop | Sort-Object SamAccountName
		#Write-Host "Native PowerShell"
	}
	
	return $UserCount
}

function Get-ADGroupObject
{
	Param (
		[Parameter(Mandatory = $true)]
		[string]$Search,
		[Parameter(Mandatory = $true)]
		[string]$FilterType		
	)
	
	if ($RemoteADModule)
	{
		$GroupCount = Invoke-Command -Session $Session -ScriptBlock {
			Param ($search,
				$filter)
			Import-Module ActiveDirectory
			Get-ADGroup -Filter { $filter -like $search } -Properties * -ErrorAction Stop | Sort-Object Name
		} -ArgumentList $Search, $FilterType
		#Write-Host "Remote PowerShell"
	}
	else
	{
		$GroupCount = Get-ADGroup -Filter { $FilterType -like $Search } -Properties * -ErrorAction Stop | Sort-Object Name
		#Write-Host "Native PowerShell"
	}
	
	return $GroupCount
}

function Get-ADGroupMemberObject
{
	Param (
		[Parameter(Mandatory = $true)]
		[string]$Search
	)
	
	if ($RemoteADModule)
	{
		$GroupMemberCount = Invoke-Command -Session $Session -ScriptBlock {
			Param ($search)
			Import-Module ActiveDirectory
			Get-ADGroupMember -Identity $search -ErrorAction Stop | Sort-Object Name
		} -ArgumentList $Search
		#Write-Host "Remote PowerShell"
	}
	else
	{
		$GroupMemberCount = Get-ADGroupMember -Identity $Search -ErrorAction Stop | Sort-Object Name
		#Write-Host "Native PowerShell"
	}
	
	return $GroupMemberCount
}

function Add-ADUserToGroup
{
	Param (
		[Parameter(Mandatory = $true)]
		[string]$GroupName,
		[Parameter(Mandatory = $true)]
		[string]$UserName
	)
	
	if ($RemoteADModule)
	{
		$CmdOutput = Invoke-Command -Session $Session -ScriptBlock {
			param ($group,
				$user)
			$VerbosePreference = 'Continue'
			Import-Module ActiveDirectory
			try
			{
				Add-ADGroupMember -Identity $group -Members $user -ErrorAction Stop
				return $true
			}
			catch
			{
				return $_
			}
		} -ArgumentList $GroupName, $UserName
	}
	else
	{
		try
		{
			Add-ADGroupMember -Identity $GroupName -Members $UserName -ErrorAction Stop
			$CmdOutput = $true
		}
		catch
		{
			$CmdOutput = $_
		}
	}
	
	return $CmdOutput
}

function Remove-ADUserFromGroup
{
	Param (
		[Parameter(Mandatory = $true)]
		[string]$GroupName,
		[Parameter(Mandatory = $true)]
		[string]$UserName
	)
	
	if ($RemoteADModule)
	{
		$CmdOutput = Invoke-Command -Session $Session -ScriptBlock {
			param ($group,
				$user)
			$VerbosePreference = 'Continue'
			Import-Module ActiveDirectory
			try
			{
				Remove-ADGroupMember -Identity $group -Members $user -Confirm:$false -ErrorAction Stop
				return $true
			}
			catch
			{
				return $_
			}
		} -ArgumentList $GroupName, $UserName
	}
	else
	{
		try
		{
			Remove-ADGroupMember -Identity $GroupName -Members $UserName -Confirm:$false -ErrorAction Stop
			$CmdOutput = $true
		}
		catch
		{
			$CmdOutput = $_
		}
	}
	
	return $CmdOutput
}

function Update-ListBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ListBox or CheckedListBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ListBox control.
	
	.PARAMETER ListBox
		The ListBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ListBox's Items collection.
	
	.PARAMETER DisplayMember
		Indicates the property to display for the items in this control.
		
	.PARAMETER ValueMember
		Indicates the property to use for the value of the control.
	
	.PARAMETER Append
		Adds the item(s) to the ListBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ListBox $ListBox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ListBox $listBox1 "Red" -Append
		Update-ListBox $listBox1 "White" -Append
		Update-ListBox $listBox1 "Blue" -Append
	
	.EXAMPLE
		Update-ListBox $listBox1 (Get-Process) "ProcessName"
	
	.NOTES
		Additional information about the function.
#>
	
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.Windows.Forms.ListBox]$ListBox,
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[Parameter(Mandatory = $false)]
		[string]$ValueMember,
		[switch]$Append
	)
	
	if (-not $Append)
	{
		$ListBox.Items.Clear()
	}
	
	if ($Items -is [System.Windows.Forms.ListBox+ObjectCollection] -or $Items -is [System.Collections.ICollection])
	{
		$ListBox.Items.AddRange($Items)
	}
	elseif ($Items -is [System.Collections.IEnumerable])
	{
		$ListBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ListBox.Items.Add($obj)
		}
		$ListBox.EndUpdate()
	}
	else
	{
		$ListBox.Items.Add($Items)
	}
	
	$ListBox.DisplayMember = $DisplayMember
	$ListBox.ValueMember = $ValueMember
}

function Show-SplashScreen
{
	<#
	.SYNOPSIS
		Displays a splash screen using the specified image.
	
	.PARAMETER Image
		Mandatory Image object that is displayed in the splash screen.
	
	.PARAMETER Title
		(Optional) Sets a title for the splash screen window. 
	
	.PARAMETER Timeout
		The amount of seconds before the splash screen is closed.
		Set to 0 to leave the splash screen open indefinitely.
		Default: 2
	
	.PARAMETER ImageLocation
		The file path or url to the image.

	.PARAMETER PassThru
		Returns the splash screen form control. Use to manually close the form.
	
	.PARAMETER Modal
		The splash screen will hold up the pipeline until it closes.

	.EXAMPLE
		PS C:\> Show-SplashScreen -Image $Image -Title 'Loading...' -Timeout 3

	.EXAMPLE
		PS C:\> Show-SplashScreen -ImageLocation 'C:\Image\MyImage.png' -Title 'Loading...' -Timeout 3

	.EXAMPLE
		PS C:\> $splashScreen = Show-SplashScreen -Image $Image -Title 'Loading...' -PassThru
				#close the splash screen
				$splashScreen.Close()
	.OUTPUTS
		System.Windows.Forms.Form
	
	.NOTES
		Created by SAPIEN Technologies, Inc.

		The size of the splash screen is dependent on the image.
		The required assemblies to use this function outside of a WinForms script:
		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing
#>
	[OutputType([System.Windows.Forms.Form])]
	param
	(
		[Parameter(ParameterSetName = 'Image',
				   Mandatory = $true,
				   Position = 1)]
		[ValidateNotNull()]
		[System.Drawing.Image]$Image,
		[Parameter(Mandatory = $false)]
		[string]$Title,
		[int]$Timeout = 2,
		[Parameter(ParameterSetName = 'ImageLocation',
				   Mandatory = $true,
				   Position = 1)]
		[ValidateNotNullOrEmpty()]
		[string]$ImageLocation,
		[switch]$PassThru,
		[switch]$Modal
	)
	
	#Create a splash screen form to display the image.
	$splashForm = New-Object System.Windows.Forms.Form
	
	#Create a picture box for the image
	$pict = New-Object System.Windows.Forms.PictureBox
	
	if ($Image)
	{
		$pict.Image = $Image;
	}
	else
	{
		$pict.Load($ImageLocation)
	}
	
	$pict.AutoSize = $true
	$pict.Dock = 'Fill'
	$splashForm.Controls.Add($pict)
	
	#Display a title if defined.
	if ($Title)
	{
		$splashForm.Text = $Title
		$splashForm.FormBorderStyle = 'FixedDialog'
	}
	else
	{
		$splashForm.FormBorderStyle = 'None'
	}
	
	#Set a timer
	if ($Timeout -gt 0)
	{
		$timer = New-Object System.Windows.Forms.Timer
		$timer.Interval = $Timeout * 1000
		$timer.Tag = $splashForm
		$timer.add_Tick({
				$this.Tag.Close();
				$this.Stop()
			})
		$timer.Start()
	}
	
	#Show the form
	$splashForm.AutoSize = $true
	$splashForm.AutoSizeMode = 'GrowAndShrink'
	$splashForm.ControlBox = $false
	$splashForm.StartPosition = 'CenterScreen'
	$splashForm.TopMost = $true
	
	if ($Modal) { $splashForm.ShowDialog() }
	else { $splashForm.Show() }
	
	if ($PassThru)
	{
		return $splashForm
	}
}
#endregion

#region Variables
[string]$ScriptDirectory = Get-ScriptDirectory
$Module = "ActiveDirectory"
$DomainControllers = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().DomainControllers.Name
$AvailableLanguages = @("English","Nederlands","Deutsch")
$VersionNumber = [System.Windows.Forms.Application]::ProductVersion
$OrgName = "Radboudumc $([char]0x00A9)"
$Object = @()
#endregion

#region Security
#$pwd = "torMVNCdF8f6uQHPD7rF"
#$passwrd = $Pwd | ConvertTo-SecureString -AsPlainText -Force
#$cred = New-Object System.Management.Automation.PsCredential("UMCN\SVC10286", $passwrd)
#endregion

#region Font Styles
$FontRegular = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Regular)
$FontBold = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold)
$FontItalic = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Italic)
$FontBoldItalic = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold + [System.Drawing.FontStyle]::Italic)
$FontStrikeout = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Strikeout)
$FontUnderline = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Strikeout)
$FontRegularStrikeout = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Regular + [System.Drawing.FontStyle]::Strikeout)
$FontRegularUnderline = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Regular + [System.Drawing.FontStyle]::Underline)
$FontRegularStrikeoutUnderline = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Regular + [System.Drawing.FontStyle]::Strikeout + [System.Drawing.FontStyle]::Underline)
$FontBoldStrikeout = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold + [System.Drawing.FontStyle]::Strikeout)
$FontBoldUnderline = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold + [System.Drawing.FontStyle]::Underline)
$FontBoldStrikeoutUnderline = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold + [System.Drawing.FontStyle]::Strikeout + [System.Drawing.FontStyle]::Underline)
$FontBoldItalicStrikeout = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold + [System.Drawing.FontStyle]::Italic + [System.Drawing.FontStyle]::Strikeout)
$FontBoldItalicUnderline = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold + [System.Drawing.FontStyle]::Italic + [System.Drawing.FontStyle]::Underline)
$FontBoldItalicStrikeoutUnderline = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Bold + [System.Drawing.FontStyle]::Italic + [System.Drawing.FontStyle]::Strikeout + [System.Drawing.FontStyle]::Underline)
$FontStrikeoutUnderline = [System.Drawing.Font]::new('Microsoft Sans Serif', 8.25, [System.Drawing.FontStyle]::Strikeout + [System.Drawing.FontStyle]::Underline)
#endregion

#region Language Culture
$UICulture = (Get-UICulture).Name
$LanguageStrings = Set-UILanguage $UICulture
#endregion
	