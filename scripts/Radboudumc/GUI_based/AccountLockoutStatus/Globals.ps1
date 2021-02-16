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
		### Dutch language
		$Language = 'Dutch'
		
		### Form Controls
		
		$FormLabel = "Accountvergrendelingsstatus v$($VersionNumber) - $($OrgName)"
		
		$GroupEventDetails = 'Event details:'
		$RadioLockout = 'Vergrendeling'
		$RadioBadPassword = 'Fout wachtwoord'
		$LabelDaysFromToday = 'Dagen sinds vandaag'
		$GroupDomainControllers = 'Domeincontrollers:'
		$CheckDomainControllers = 'Zoek specifieke domeincontroller'
		$GroupUsername = 'Gebruikersnaam:'
		$CheckUsername = 'Zoek specifieke gebruiker'
		$ButtonCheck = "&Controleer"
		$ButtonSearch = "&Zoeken"
		$ButtonAgain = "&Opnieuw"
		$ButtonExit = "&Afsluiten"
		$LabelLanguage = 'Taal:'
		
		### Tooltip text
		$TTbtnCheck = "Deze actie moet eerst worden uitgevoerd. Als er geen validatiefouten worden gevonden, wordt de knop 'Zoeken' ingeschakeld."
		$TTbtnSearch = 'Deze actie zal de zoekactie starten op basis van de geselecteerde opties.'
		$TTbtnAgain = 'Deze actie zal het formulier terugzetten naar de oorspronkelijke staat.'
		$TTbtnExit = 'Deze actie zal het formulier sluiten.'
		$TTcmbLanguage = 'U kunt hier de formuliertaal selecteren.'
		$TTradLockout = 'Indien geselecteerd, zal er naar lockout-evenementen worden gezocht.'
		$TTradBadPassword = 'Indien geselecteerd, wordt gezocht naar gebeurtenissen met een ongeldig wachtwoord.'
		$TTnumDaysFromToday = 'Specificeert hoeveel dagen vanaf vandaag het begin van de zoekopdracht zal zijn.'
		$TTchkDomainController = "Indien aangevinkt, kunnen specifieke domeincontrollers worden geselecteerd.`r`nIndien niet aangevinkt, zullen alle domeincontrollers worden gebruikt."
		$TTlstDomainController = "Maak een lijst van de beschikbare domeincontrollers om op te zoeken.`r`nMultiselect is mogelijk met Shift en Ctrl."
		$TTchkUsername = 'Indien aangevinkt, zal de zoekopdracht gebaseerd zijn op een specifieke gebruiker.'
		$TTtxtUsername = 'Vul dit veld in met de gebruikersnaam.'
		$TTrtbOutput = 'Dit is het logscherm.'
		
		### Verbose output
		$VBInfo = 'INFO:'
		$VBWarning = 'WAARSCHUWING:'
		$VBError = 'FOUT:'
		$VBSuccess = 'SUCCES:'
		$VBModuleCheck = 'Bezig met zoeken naar'
		$VBModule = 'module'
		$VBModuleAlreadyImported = 'module is al geïmporteerd.'
		$VBModuleAvailable = 'module beschikbaar. Importeren ... '
		$VBModuleImported = 'module geïmporteerd.'
		$VBModuleSearch = 'Zoeken naar'
		$VBModuleOnline = 'module in online galerij.'
		$VBModuleFound = 'module gevonden. Installeren ... '
		$VBModuleInstalled = 'module geïnstalleerd. Importeren ... '
		$VBModuleFailed = 'module niet geïmporteerd, niet beschikbaar en niet gevonden in online galerij.'
		$VBPSSessionSetup = 'PS externe sessie aan het opzetten.'
		$VBPSSessionSetupSuccess = 'PS externe verbinding tot stand gebracht.'
		$VBPSSessionSetupFailed = 'PS remoting kon geen verbinding maken, neem contact op met uw systeembeheerder.'
		$VBPSSessionImport = 'Externe PS-sessie importeren.'
		$VBPSSessionImportSuccess = 'PS externe sessie geïmporteerd.'
		$VBPSSessionImportFailed = 'PS remoting kon sessie niet importeren. Neem contact op met uw systeembeheerder om door te gaan. '
		$VBPSSessionInvalid = 'PS remoting geen geldige sessie gevonden. Neem contact op met uw systeembeheerder om door te gaan. '

		$VBNoDCSelected = 'Selecteer een of meer domeincontrollers waarop u een zoekopdracht wilt uitvoeren.'
		$VBNoUsername = 'Vul een gebruikersnaam in'
		$VBSearchParam = 'Zoekparameters:'
		$VBEventType = 'Gebeurtenistype:'
		$VBLockout = 'Vergrendeling'
		$VBBadPassword = 'Slecht wachtwoord'
		$VBDaysFromToday = "$($LabelDaysFromToday):"
		$VBUsername = $GroupUsername
		$VBAllUsers = 'Alle gebruikers'
		$VBDomainControllers = $GroupDomainControllers
		$VBMoreDCSelected = 'Meer dan 1 domeincontroller geselecteerd. Dit kan veel tijd kosten. Wees alstublieft geduldig.'
		$VBMoreNumDays = 'Langer dan 1 dag zoeken in het verleden. Dit kan veel tijd kosten. Wees alstublieft geduldig.'

		$VBSearchOn = 'Zoekopdracht uitvoeren op:'
		$VBGetEventError = 'Kan events niet ophalen.'
		$VBEventsProcessed = 'Verwerkte events:'
		$VBOutput = 'Uitvoer:'
		$VBNoLockoutsUser = 'Geen uitsluitingen gevonden voor gebruiker'
		$VBNoBadPasswordUser = 'Geen slechte wachtwoorden gevonden voor gebruiker'
		$VBNoLockouts = 'Geen uitsluitingen gevonden'
		$VBNoBadPassword = 'Geen slechte wachtwoordpogingen gevonden'

		### Splash Screen
		$SSTitle = 'Bezig met laden...'
		$SSTitleEventSearch = 'Events zoeken...'
	}
	elseif ($Culture -like "de-*")
	{
		### German Language
		$Language = 'Englisch'
		
		### Form Controls
		$FormLabel = "Kontosperrungsstatus v$($VersionNumber) - $($OrgName)"
		
		$GroupEventDetails = 'Ereignisdetails:'
		$RadioLockout = 'Aussperrung'
		$RadioBadPassword = 'Ungültiges Passwort'
		$LabelDaysFromToday = 'Tage ab heute'
		$GroupDomainControllers = 'Domänencontroller:'
		$CheckDomainControllers = 'Suchspezifischen Domänencontroller'
		$GroupUsername = 'Benutzername:'
		$CheckUsername = 'Bestimmten Benutzer suchen'
		$ButtonCheck = "&Check"
		$ButtonSearch = "&Suche"
		$ButtonAgain = "&Nochmal"
		$ButtonExit = "&Ausschalten"
		$LabelLanguage = 'Sprache:'
		
		### Tooltip Text
		$TTbtnCheck = "Diese Aktion muss zuerst ausgeführt werden. Wenn keine Validierungsfehler gefunden werden, wird die Schaltfläche `'Suchen`" aktiviert."
		$TTbtnSearch = 'Diese Aktion startet die Suche basierend auf den ausgewählten Optionen.'
		$TTbtnAgain = 'Diese Aktion setzt das Formular auf seinen Ausgangszustand zurück.'
		$TTbtnExit = 'Diese Aktion schließt das Formular.'
		$TTcmbLanguage = 'Hier können Sie die Formularsprache auswählen.'
		$TTradLockout = 'Wenn diese Option ausgewählt ist, wird nach Sperrereignissen gesucht.'
		$TTradBadPassword = 'Wenn diese Option ausgewählt ist, wird nach falschen Kennwortereignissen gesucht.'
		$TTnumDaysFromToday = 'Gibt an, wie viele Tage ab heute der Beginn der Suche sein wird.'
		$TTchkDomainController = 'Wenn diese Option aktiviert ist, können bestimmte Domänencontroller ausgewählt werden .`r`nWenn diese Option nicht aktiviert ist, werden alle Domänencontroller verwendet.'
		$TTlstDomainController = "Listet die verfügbaren Domänencontroller auf, nach denen gesucht werden soll .`r`nMultiselect ist mit Shift und Ctrl möglich."
		$TTchkUsername = 'Wenn diese Option aktiviert ist, basiert die Suche auf einem bestimmten Benutzer.'
		$TTtxtUsername = 'Füllen Sie dieses Feld mit dem Benutzernamen.'
		$TTrtbOutput = 'Dies ist der Protokollbildschirm.'
		
		### Verbose output
		$VBInfo = 'INFO:'
		$VBWarning = 'WARNUNG:'
		$VBError = 'FEHLER:'
		$VBSuccess = 'ERFOLG:'
		$VBModuleCheck = 'Prüfen'
		$VBModule = 'Modul'
		$VBModuleAlreadyImported = 'Modul ist bereits importiert.'
		$VBModuleAvailable = 'Modul verfügbar. Importieren ... '
		$VBModuleImported = 'Modul importiert.'
		$VBModuleSearch = 'Suchen nach'
		$VBModuleOnline = 'Modul in der Online-Galerie.'
		$VBModuleFound = 'Modul gefunden. Installieren ... '
		$VBModuleInstalled = 'Modul installiert. Importieren ... '
		$VBModuleFailed = 'Modul nicht importiert, nicht verfügbar und nicht in der Online-Galerie gefunden.'
		$VBPSSessionSetup = 'PS-Remoting-Sitzung einrichten.'
		$VBPSSessionSetupSuccess = 'PS-Remoting-Verbindung hergestellt.'
		$VBPSSessionSetupFailed = 'PS-Remoting konnte keine Verbindung herstellen. Wenden Sie sich an Ihren Systemadministrator.'
		$VBPSSessionImport = 'Remote-PS-Sitzung importieren.'
		$VBPSSessionImportSuccess = 'PS-Remoting-Sitzung importiert.'
		$VBPSSessionImportFailed = 'PS-Remoting konnte Sitzung nicht importieren. Wenden Sie sich an Ihren Systemadministrator. '
		$VBPSSessionInvalid = 'PS-Remoting keine gültige Sitzung gefunden. Wenden Sie sich an Ihren Systemadministrator. '
		
		$VBNoDCSelected = 'Bitte wählen Sie einen oder mehrere Domänencontroller aus, auf denen Sie eine Suche durchführen möchten.'
		$VBNoUsername = 'Bitte geben Sie einen Benutzernamen ein'
		$VBSearchParam = 'Suchparameter:'
		$VBEventType = 'Ereignistyp:'
		$VBLockout = 'Aussperrung'
		$VBBadPassword = 'Ungültiges Passwort'
		$VBDaysFromToday = "$($LabelDaysFromToday):"
		$VBUsername = $GroupUsername
		$VBAllUsers = 'Alle Benutzer'
		$VBDomainControllers = $GroupDomainControllers
		$VBMoreDCSelected = 'Mehr als 1 Domänencontroller ausgewählt. Dies kann viel Zeit in Anspruch nehmen. Bitte haben Sie Geduld.'
		$VBMoreNumDays = 'Suche nach mehr als einem Tag in der Vergangenheit. Dies kann viel Zeit in Anspruch nehmen. Bitte haben Sie Geduld.'
		
		$VBSearchOn = 'Suche durchführen für:'
		$VBGetEventError = 'Ereignisse konnten nicht abgerufen werden.'
		$VBEventsProcessed = 'Verarbeitete Ereignisse:'
		$VBOutput = 'Ausgabe:'
		$VBNoLockoutsUser = 'Keine Sperren für Benutzer gefunden'
		$VBNoBadPasswordUser = 'Keine falschen Passwörter für Benutzer gefunden'
		$VBNoLockouts = 'Keine Sperren gefunden'
		$VBNoBadPassword = 'Keine falschen Passwortversuche gefunden'
		
		### Splash Screen
		$SSTitle = 'Wird geladen ...'
		$SSTitleEventSearch = 'Ereignisse suchen...'
	}
	else
	{
		### English language
		$Language = 'English'
		
		### Form Controls
		$FormLabel = "Account Lockout Status v$($VersionNumber) - $($OrgName)"
		
		$GroupEventDetails = 'Event Details:'
		$RadioLockout = 'Lockout'
		$RadioBadPassword = 'Bad Password'
		$LabelDaysFromToday = 'Days from today'
		$GroupDomainControllers = 'Domain Controllers:'
		$CheckDomainControllers = 'Search specific Domain Controller'
		$GroupUsername = 'Username:'
		$CheckUsername = 'Search specific user'
		$ButtonCheck = "&Check"
		$ButtonSearch = "&Search"
		$ButtonAgain = "A&gain"
		$ButtonExit = "E&xit"
		$LabelLanguage = 'Language:'
				
		### Tooltip text
		$TTbtnCheck = "This action needs to be run first. If no validation errors are found, the `'Search`' button will be enabled."
		$TTbtnSearch = 'This action will start the search based on the selected options.'
		$TTbtnAgain = 'This action will reset the form to its initial state.'
		$TTbtnExit = 'This action will close the form.'
		$TTcmbLanguage = 'You can select the form language here.'
		$TTradLockout = 'When selected, lockout events will be searched for.'
		$TTradBadPassword = 'When selected, bad password events will be searched for.'
		$TTnumDaysFromToday = 'Specifies how many day from today the start of the search will be.'
		$TTchkDomainController = 'When checked, specific domain controllers can be selected.`r`nWhen unchecked, all domain controllers will be used.'
		$TTlstDomainController = "List the available Domain Controllers to search on.`r`nMultiselect is possible with Shift and Ctrl."
		$TTchkUsername = 'When checked, the search will be based on a specific user.'
		$TTtxtUsername = 'Populate this field with the username.'
		$TTrtbOutput = 'This is the log screen.'
				
		### Verbose output strings
		$VBInfo = 'INFO:'
		$VBWarning = 'WARNING:'
		$VBError = 'ERROR:'
		$VBSuccess = 'SUCCESS:'
		$VBModuleCheck = 'Checking for'
		$VBModule = 'module'
		$VBModuleAlreadyImported = 'module is already imported.'
		$VBModuleAvailable = 'module available. Importing...'
		$VBModuleImported = 'module imported.'
		$VBModuleSearch = 'Searching for'
		$VBModuleOnline = 'module in online gallery.'
		$VBModuleFound = 'module found. Installing...'
		$VBModuleInstalled = 'module installed. Importing...'
		$VBModuleFailed = 'module not imported, not available and not found in online gallery.'
		$VBPSSessionSetup = 'Setting up PS remoting session.'
		$VBPSSessionSetupSuccess = 'PS remoting connection established.'
		$VBPSSessionSetupFailed = 'PS remoting failed to connect, please contact your Systems Administrator.'
		$VBPSSessionImport = 'Importing remote PS session.'
		$VBPSSessionImportSuccess = 'PS remoting session imported.'
		$VBPSSessionImportFailed = 'PS remoting failed to import session. Unable to continue, please contact your Systems Administrator.'
		$VBPSSessionInvalid = 'PS remoting no valid session found. Unable to continue, please contact your Systems Administrator.'
		
		$VBNoDCSelected = 'Please select one or more Domain Controllers on which you want to perform a search.'
		$VBNoUsername = 'Please fill in a username'
		$VBSearchParam = 'Search parameters:'
		$VBEventType = 'Event type:'
		$VBLockout = 'Lockout'
		$VBBadPassword = 'Bad Password'
		$VBDaysFromToday = "$($LabelDaysFromToday):"
		$VBUsername = $GroupUsername
		$VBAllUsers = 'All users'
		$VBDomainControllers = $GroupDomainControllers
		$VBMoreDCSelected = 'More than 1 domain controller selected. This may take a considerable amount of time. Please be patient.'
		$VBMoreNumDays = 'Searching for longer than 1 day in the past. This may take a considerable amount of time. Please be patient.'
		
		$VBSearchOn = 'Performing search on:'
		$VBGetEventError = 'Failed to get events.'
		$VBEventsProcessed = 'Processed events:'
		$VBOutput = 'Output:'
		$VBNoLockoutsUser = 'No lockouts found for user '
		$VBNoBadPasswordUser = 'No bad passwords found for user '
		$VBNoLockouts = 'No lockouts found'
		$VBNoBadPassword = 'No bad password attempts found'
		
		### Splash screen
		$SSTitle = 'Loading...'
		$SSTitleEventSearch = 'Searching events...'
		$SSTitleEventProcessing = 'Processing events...'
	}
	
	$LanguageObject = @()
	
	$Properties = @{
		Language				 = $Language
		FormLabel			     = $FormLabel
		GroupEventDetails	     = $GroupEventDetails
		RadioLockout			 = $RadioLockout
		RadioBadPassword		 = $RadioBadPassword
		LabelDaysFromToday	     = $LabelDaysFromToday
		GroupDomainControllers   = $GroupDomainControllers
		CheckDomainControllers   = $CheckDomainControllers
		GroupUsername		     = $GroupUsername
		CheckUsername		     = $CheckUsername
		ButtonCheck			     = $ButtonCheck
		ButtonSearch			 = $ButtonSearch
		ButtonAgain			     = $ButtonAgain
		ButtonExit			     = $ButtonExit
		LabelLanguage		     = $LabelLanguage
		TTbtnCheck			     = $TTbtnCheck
		TTbtnSearch			     = $TTbtnSearch
		TTbtnAgain			     = $TTbtnAgain
		TTbtnExit			     = $TTbtnExit
		TTcmbLanguage		     = $TTcmbLanguage
		TTradLockout			 = $TTradLockout
		TTradBadPassword		 = $TTradBadPassword
		TTnumDaysFromToday	     = $TTnumDaysFromToday
		TTchkDomainController    = $TTchkDomainController
		TTlstDomainController    = $TTlstDomainController
		TTchkUsername		     = $TTchkUsername
		TTtxtUsername		     = $TTtxtUsername
		TTrtbOutput			     = $TTrtbOutput
		VBInfo				     = $VBInfo
		VBWarning			     = $VBWarning
		VBError				     = $VBError
		VBSuccess			     = $VBSuccess
		VBModuleCheck		     = $VBModuleCheck
		VBModule				 = $VBModule
		VBModuleAlreadyImported  = $VBModuleAlreadyImported
		VBModuleAvailable	     = $VBModuleAvailable
		VBModuleImported		 = $VBModuleImported
		VBModuleSearch		     = $VBModuleSearch
		VBModuleOnline		     = $VBModuleOnline
		VBModuleFound		     = $VBModuleFound
		VBModuleInstalled	     = $VBModuleInstalled
		VBModuleFailed		     = $VBModuleFailed
		VBPSSessionSetup		 = $VBPSSessionSetup
		VBPSSessionSetupSuccess  = $VBPSSessionSetupSuccess
		VBPSSessionSetupFailed   = $VBPSSessionSetupFailed
		VBPSSessionImport	     = $VBPSSessionImport
		VBPSSessionImportSuccess = $VBPSSessionImportSuccess
		VBPSSessionImportFailed  = $VBPSSessionImportFailed
		VBPSSessionInvalid	     = $VBPSSessionInvalid
		VBNoDCSelected		     = $VBNoDCSelected
		VBNoUsername			 = $VBNoUsername
		VBSearchParam		     = $VBSearchParam
		VBEventType			     = $VBEventType
		VBLockout			     = $VBLockout
		VBBadPassword		     = $VBBadPassword
		VBDaysFromToday		     = $VBDaysFromToday
		VBUsername			     = $VBUsername
		VBAllUsers			     = $VBAllUsers
		VBDomainControllers	     = $VBDomainControllers
		VBMoreDCSelected		 = $VBMoreDCSelected
		VBMoreNumDays		     = $VBMoreNumDays
		VBSearchOn			     = $VBSearchOn
		VBGetEventError		     = $VBGetEventError
		VBEventsProcessed	     = $VBEventsProcessed
		VBOutput				 = $VBOutput
		VBNoLockoutsUser		 = $VBNoLockoutsUser
		VBNoBadPasswordUser	     = $VBNoBadPasswordUser
		VBNoLockouts			 = $VBNoLockouts
		VBNoBadPassword		     = $VBNoBadPassword
		SSTitle				     = $SSTitle
		SSTitleEventSearch	     = $SSTitleEventSearch
	}
	
	$LanguageObject = New-Object -TypeName PSObject -Property $Properties | Select-Object *
	
	return $LanguageObject
}
#endregion

#region Variables
[string]$ScriptDirectory = Get-ScriptDirectory
$Module = "ActiveDirectory"
$AllDomainControllers = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().DomainControllers.Name | Sort-Object
$AvailableLanguages = @("English","Nederlands","Deutsch")
$VersionNumber = [System.Windows.Forms.Application]::ProductVersion
$OrgName = "Radboudumc $([char]0x00A9)"
#endregion

#region Security
$pwd = "torMVNCdF8f6uQHPD7rF"
$Object = @()
$passwrd = $Pwd | ConvertTo-SecureString -AsPlainText -Force
$cred = New-Object System.Management.Automation.PsCredential("UMCN\SVC10286", $passwrd)
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
	