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

function Build-RemoteExchange2013Session
{
	###################### Build Remote PSSession  ############################################################################
	if (get-pssession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" })
	{
		return 0
	}
	
	$CASServers = @("UMCEXCHP01", "UMCEXCHP02", "UUMCEXCHP03", "UMCEXCHP04", "UMCEXCHP05", "UMCEXCHP06")
	
	foreach ($script:CASServer in $CASServers)
	{
		$found = Test-Connection $script:CASServer -Count 2 -Delay 1 -ErrorAction SilentlyContinue
		if ($found) { Break }
	}
	
	if (!$found)
	{
		return 1
	}
	
	$Exchange2013Server = $script:CASServer
	
	#Try implicit remoting-only works for Exchange 2013
	try
	{
		$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$Exchange2013Server/PowerShell/ -Authentication Kerberos -Credential $cred -ErrorAction Stop
		Import-PSSession $ExchangeSession -ErrorAction Stop -verbose:$false
		return 2
	}
	catch
	{
		return 3
	}
	### End Remote PSSession Build
}

function Set-UILanguage ($Culture)
{
	if ($Culture -like "nl-*")
	{
		### Dutch language
		$Language = 'Dutch'
		
		### Form Controls
		$FormLabel = "Blokkeer spam`/phishing-e-mails v$($VersionNumber) - $($OrgName)"
		
		$ButtonCheck = "&Controleer"
		$ButtonBlock = "&Blokkeren"
		$ButtonAgain = "&Opnieuw"
		$ButtonExit = "&Afsluiten"
		$CheckBlockEmailAddress = 'Blokkeren op e-mailadres'
		$CheckBlockSubjectBody = 'Blokkeer op trefwoorden in BODY of SUBJECT'
		$LabelVerboseOutput = 'Uitgebreide uitvoer:'
		$LabelLanguage = 'Taal:'
		$GroupParameters = 'Parameters'
		
		### Tooltip-tekst
		$TTbtnCheck = "Deze actie moet eerst worden uitgevoerd. Als er geen validatiefouten worden gevonden, wordt de knop `'Blokkeren`' ingeschakeld."
		$TTbtnBlock = 'Deze actie zal de corresponderende mailflow regels updaten, gebaseerd op de parameters.'
		$TTbtnAgain = 'Deze actie zal het formulier terugzetten naar de beginstaat.'
		$TTbtnExit = 'Deze actie zal het formulier sluiten.'
		$TTcmbLanguage = 'U kunt hier de formuliertaal selecteren.'
		$TTtxtEmailAddress = 'Vul dit veld met het e-mailadres dat u wilt blokkeren.'
		$TTtxtSubjectBody = 'Vul dit veld met de trefwoorden die u wilt blokkeren als ze in het onderwerp of de hoofdtekst worden gevonden.'
		$TTchkBlockEmailAddress = 'Vink dit vakje aan om een ​​specifiek e-mailadres te kunnen blokkeren.'
		$TTchkBlockSubjectBody = 'Vink dit vakje aan om trefwoorden te blokkeren die in het onderwerp of de hoofdtekst worden gevonden.'
		$TTrtbOutput = "Dit is het logscherm."

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
		$VBSessionAlreadySetup = 'PS-remoting sessie is al aanwezig.'
		$VBPSSessionSetup = 'PS-remoting sessie aan het opzetten.'
		$VBPSSessionSetupSuccess = 'PS-remoting verbinding tot stand gebracht.'
		$VBPSSessionSetupFailed = 'PS-remoting kon geen verbinding maken, neem contact op met uw systeembeheerder.'
		$VBPSSessionImport = 'PS-remoting sessie importeren.'
		$VBPSSessionImportSuccess = 'PS-remoting sessie geïmporteerd.'
		$VBPSSessionImportFailed = 'PS-remoting kon sessie niet importeren. Neem contact op met uw systeembeheerder om door te gaan.'
		$VBPSSessionInvalid = 'PS-remoting geen geldige sessie gevonden. Neem contact op met uw systeembeheerder om door te gaan.'
		$VBSessionNoCAS = 'Geen CAS server gevonden. Neem contact op met uw systeembeheerder.'
		
		$VBNoEmailAddress = 'Vul een e-mailadres in.'
		$VBPrevCountEmail = 'Aantal oude geblokkeerde e-mailadressen:'
		$VBNewCountEmail = 'Aantal nieuwe geblokkeerde e-mailadressen:'
		$VBNoKeywords = 'Vul een tekstzin in om op te filteren.'
		$VBPrevCountKeywords = 'Vorig onderwerp / body block count:'
		$VBNewCountKeywords = 'Aantal nieuwe onderwerp / body block:'
		$VBNoSelection = 'Om verder te gaan, moet ten minste een van de opties worden geselecteerd.'
		$VBBlockParam = 'Blokkeer parameters:'
		$VBBlockEmail = 'E-mailadres dat moet worden geblokkeerd:'
		$VBBlockKeywords = 'Tekstzin die moet worden geblokkeerd:'
		
		$VBBlockCmd = 'Het volgende commando wordt uitgevoerd:'
		$VBBlockEmailSuccess = 'E-mailadres toegevoegd aan blokkeerlijst:'
		$VBBlockEmailFailed = 'Toevoegen van e-mailadres aan blokkeerlijst is mislukt:'
		$VBBlockKeywordsSuccess = 'Tekstzin toegevoegd aan blokkeerlijst:'
		$VBBlockKeywordsFailed = 'Toevoegen van tekstzin aan blokkeerlijst is mislukt:'

		### Splash-scherm
		$SSTitle = 'Bezig met laden ...'
	}
	elseif ($Culture -like "de-*")
	{
		### German language
		$Language = 'German'
		
		### Form Controls
		$FormLabel = "Blockieren Spam-`/Phishing-E-Mails v$($VersionNumber) - $($OrgName)"
		$ButtonCheck = "&Prüfen"
		$ButtonBlock = "&Block"
		$ButtonAgain = "&Nochmal"
		$ButtonExit = "&Ausschalten"
		$CheckBlockEmailAddress = 'E-Mail-Adresse blockieren'
		$CheckBlockSubjectBody = 'Schlüsselwörter in BODY oder SUBJECT blockieren'
		$LabelVerboseOutput = 'Ausführliche Ausgabe:'
		$LabelLanguage = 'Sprache:'
		$GroupParameters = 'Parameter'
		
		### Tooltip Text
		$TTbtnCheck = "Diese Aktion muss zuerst ausgeführt werden. Wenn keine Validierungsfehler gefunden werden, wird die Schaltfläche `'Blockieren`' aktiviert."
		$TTbtnBlock = 'Diese Aktion aktualisiert die entsprechenden Mailflow-Regeln basierend auf den Parametern.'
		$TTbtnAgain = 'Diese Aktion setzt das Formular auf den Ausgangszustand zurück.'
		$TTbtnExit = 'Diese Aktion schließt das Formular.'
		$TTcmbLanguage = 'Hier können Sie die Formularsprache auswählen.'
		$TTtxtEmailAddress = 'Füllen Sie dieses Feld mit der E-Mail-Adresse aus, die Sie blockieren möchten.'
		$TTtxtSubjectBody = 'Füllen Sie dieses Feld mit den Schlüsselwörtern aus, die Sie blockieren möchten, wenn sie im Betreff oder im Text gefunden werden.'
		$TTchkBlockEmailAddress = 'Aktivieren Sie dieses Kontrollkästchen, um eine bestimmte E-Mail-Adresse blockieren zu können.'
		$TTchkBlockSubjectBody = 'Aktivieren Sie dieses Kontrollkästchen, um Schlüsselwörter zu blockieren, die im Betreff oder im Text gefunden wurden.'
		$TTrtbOutput = "Dies ist der Protokollbildschirm."
		
		### Verbose output
		$VBInfo = 'INFO:'
		$VBWarning = 'WARNUNG:'
		$VBError = 'FEHLER:'
		$VBSuccess = 'ERFOLG:'
		$VBModuleCheck = 'Prüfen auf'
		$VBModule = 'Modul'
		$VBModuleAlreadyImported = 'Modul ist bereits importiert.'
		$VBModuleAvailable = 'Modul verfügbar. Importieren ... '
		$VBModuleImported = 'Modul importiert.'
		$VBModuleSearch = 'Suchen nach'
		$VBModuleOnline = 'Modul in der Online-Galerie.'
		$VBModuleFound = 'Modul gefunden. Installieren ... '
		$VBModuleInstalled = 'Modul installiert. Importieren ... '
		$VBModuleFailed = 'Modul nicht importiert, nicht verfügbar und nicht in der Online-Galerie gefunden.'
		$VBSessionAlreadySetup = 'PS-remoting-Sitzung ist bereits vorhanden.'
		$VBPSSessionSetup = 'PS-Remoting-Sitzung einrichten.'
		$VBPSSessionSetupSuccess = 'PS-Remoting-Verbindung hergestellt.'
		$VBPSSessionSetupFailed = 'PS-Remoting konnte keine Verbindung herstellen. Wenden Sie sich an Ihren Systemadministrator.'
		$VBPSSessionImport = 'Remote-PS-Sitzung importieren.'
		$VBPSSessionImportSuccess = 'PS-Remoting-Sitzung importiert.'
		$VBPSSessionImportFailed = 'PS-Remoting konnte Sitzung nicht importieren. Wenden Sie sich an Ihren Systemadministrator.'
		$VBPSSessionInvalid = 'PS-Remoting keine gültige Sitzung gefunden. Wenden Sie sich an Ihren Systemadministrator.'
		$VBSessionNoCAS = 'Keine CAS-Server gefunden. Wenden Sie sich an Ihren Systemsadministrator.'
		
		$VBNoEmailAddress = 'Bitte geben Sie eine E-Mail-Adresse ein.'
		$VBPrevCountEmail = 'Anzahl der vorherigen E-Mail-Adressblöcke:'
		$VBNewCountEmail = 'Neue Anzahl von E-Mail-Adressblöcken:'
		$VBNoKeywords = 'Bitte geben Sie eine Textphrase ein, nach der gefiltert werden soll.'
		$VBPrevCountKeywords = 'Anzahl der vorherigen Betreff-/Körperblöcke:'
		$VBNewCountKeywords = 'Anzahl neuer Betreff-/Körperblöcke:'
		$VBNoSelection = 'Um fortzufahren, muss mindestens eine der Optionen ausgewählt werden.'
		$VBBlockParam = 'Blockparameter:'
		$VBBlockEmail = 'Zu blockierende E-Mail-Adresse:'
		$VBBlockKeywords = 'Zu blockierende Textphrase:'
		
		$VBBlockCmd = 'Den folgenden Befehl wird ausgeführt:'
		$VBBlockEmailSuccess = 'E-Mail-Adresse zur Sperrliste hinzugefügt:'
		$VBBlockEmailFailed = 'E-Mail-Adresse konnte nicht zur Sperrliste hinzugefügt werden:'
		$VBBlockKeywordsSuccess = 'Textphrase zur Sperrliste hinzugefügt:'
		$VBBlockKeywordsFailed = 'Fehler beim Hinzufügen einer Textphrase zur Blockliste:'
		
		### Splash Screen
		$SSTitle = 'Wird geladen...'
	}
	else
	{
		### English language
		$Language = 'English'
		
		### Form Controls
		$FormLabel = "Block Spam`/Phishing Emails v$($VersionNumber) - $($OrgName)"
		
		$ButtonCheck = "&Check"
		$ButtonBlock = "&Block"
		$ButtonAgain = "A&gain"
		$ButtonExit = "E&xit"
		$CheckBlockEmailAddress = 'Block on emailaddress'
		$CheckBlockSubjectBody = 'Block on keywords in BODY or SUBJECT'
		$LabelVerboseOutput = 'Verbose output:'
		$LabelLanguage = 'Language:'
		$GroupParameters = 'Parameters'
		
		### Tooltip text
		$TTbtnCheck = "This action needs to be run first. If no validation errors are found, the `'Block`' button will be enabled."
		$TTbtnBlock = 'This action will update the corresponding mailflow rules, based on the parameters.'
		$TTbtnAgain = 'This action will reset the form to its initial state.'
		$TTbtnExit = 'This action will close the form.'
		$TTcmbLanguage = 'You can select the form language here.'
		$TTtxtEmailAddress = 'Populate this field with the emailaddress you wish to block.'
		$TTtxtSubjectBody = 'Populate this field with the keywords you wish to block if found in the subject or body.'
		$TTchkBlockEmailAddress = 'Check this box to be able to block a specific emailaddress.'
		$TTchkBlockSubjectBody = 'Check this box to be abke to block keywords found in the subject or body.'
		$TTrtbOutput = "This is the log screen."
				
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
		$VBSessionAlreadySetup = 'PS-remoting session already present.'
		$VBPSSessionSetup = 'Setting up PS remoting session.'
		$VBPSSessionSetupSuccess = 'PS remoting connection established.'
		$VBPSSessionSetupFailed = 'PS remoting failed to connect, please contact your Systems Administrator.'
		$VBPSSessionImport = 'Importing remote PS session.'
		$VBPSSessionImportSuccess = 'PS remoting session imported.'
		$VBPSSessionImportFailed = 'PS remoting failed to import session. Unable to continue, please contact your Systems Administrator.'
		$VBPSSessionInvalid = 'PS remoting no valid session found. Unable to continue, please contact your Systems Administrator.'
		$VBSessionNoCAS = 'No CAS server found. Unable to continue, please contact your Systems Administrator.'
		$VBNoEmailAddress = 'Please fill in an emailaddress.'
		$VBPrevCountEmail = 'Previous emailaddress block count:'
		$VBNewCountEmail = 'New emailaddress block count:'
		$VBNoKeywords = 'Please fill in a text phrase to filter on.'
		$VBPrevCountKeywords = 'Previous subject/body block count:'
		$VBNewCountKeywords = 'New subject/body block count:'
		$VBNoSelection = 'In order to proceed, at least one of the options need to be selected.'
		$VBBlockParam = 'Block parameters:'
		$VBBlockEmail = 'Emailaddress to be blocked:'
		$VBBlockKeywords = 'Text phrase to be blocked:'
		
		$VBBlockCmd = 'Performing the following command:'
		$VBBlockEmailSuccess = 'Added emailaddress to block list:'
		$VBBlockEmailFailed = 'Failed to add emailaddress to block list:'
		$VBBlockKeywordsSuccess = 'Added text phrase to block list:'
		$VBBlockKeywordsFailed = 'Failed to add text phrase to block list:'
		
		
		### Splash screen
		$SSTitle = 'Loading...'
	}
	
	$LanguageObject = @()
	
	$Properties = @{
		Language			    	= $Language
		FormLabel			    	= $FormLabel
		ButtonCheck			    	= $ButtonCheck
		ButtonBlock			    	= $ButtonBlock
		ButtonAgain					= $ButtonAgain
		ButtonExit			    	= $ButtonExit
		LabelVerboseOutput	    	= $LabelVerboseOutput
		LabelLanguage		    	= $LabelLanguage
		GroupParameters		    	= $GroupParameters
		CheckBlockEmailAddress		= $CheckBlockEmailAddress
		CheckBlockSubjectBody		= $CheckBlockSubjectBody
		TTtxtEmailAddress	    	= $TTtxtEmailAddress
		TTtxtSubjectBody			= $TTtxtSubjectBody
		TTbtnCheck			    	= $TTbtnCheck
		TTbtnBlock			      	= $TTbtnBlock
		TTbtnAgain					= $TTbtnAgain
		TTbtnExit			    	= $TTbtnExit
		TTchkBlockEmailAddress  	= $TTchkBlockEmailAddress
		TTchkBlockSubjectBody		= $TTchkBlockSubjectBody
		TTrtbOutput			     	= $TTrtbOutput
		SSTitle				    	= $SSTitle
		VBInfo				    	= $VBInfo
		VBWarning			    	= $VBWarning
		VBError				    	= $VBError
		VBSuccess					= $VBSuccess
		VBModuleCheck		     	= $VBModuleCheck
		VBModule				 	= $VBModule
		VBModuleAlreadyImported  	= $VBModuleAlreadyImported
		VBModuleAvailable	     	= $VBModuleAvailable
		VBModuleImported		 	= $VBModuleImported
		VBModuleSearch		     	= $VBModuleSearch
		VBModuleOnline		     	= $VBModuleOnline
		VBModuleFound		     	= $VBModuleFound
		VBModuleInstalled	     	= $VBModuleInstalled
		VBModuleFailed		     	= $VBModuleFailed
		VBSessionAlreadySetup		= $VBSessionAlreadySetup
		VBPSSessionSetup		 	= $VBPSSessionSetup
		VBPSSessionSetupSuccess  	= $VBPSSessionSetupSuccess
		VBPSSessionSetupFailed   	= $VBPSSessionSetupFailed
		VBPSSessionImport	     	= $VBPSSessionImport
		VBPSSessionImportSuccess 	= $VBPSSessionImportSuccess
		VBPSSessionImportFailed  	= $VBPSSessionImportFailed
		VBPSSessionInvalid	    	= $VBPSSessionInvalid
		VBSessionNoCAS		    	= $VBSessionNoCAS
		VBNoEmailAddress	     	= $VBNoEmailAddress
		VBPrevCountEmail	     	= $VBPrevCountEmail
		VBNewCountEmail		 		= $VBNewCountEmail
		VBNoKeywords		     	= $VBNoKeywords
		VBPrevCountKeywords	 		= $VBPrevCountKeywords
		VBNewCountKeywords	     	= $VBNewCountKeywords
		VBNoSelection		     	= $VBNoSelection
		VBBlockParam		     	= $VBBlockParam
		VBBlockEmail		     	= $VBBlockEmail
		VBBlockKeywords		  	   	= $VBBlockKeywords
		VBBlockCmd			     	= $VBBlockCmd
		VBBlockEmailSuccess	 		= $VBBlockEmailSuccess
		VBBlockEmailFailed	     	= $VBBlockEmailFailed
		VBBlockKeywordsSuccess  	= $VBBlockKeywordsSuccess
		VBBlockKeywordsFailed		= $VBBlockKeywordsFailed
	}
	
	$LanguageObject = New-Object -TypeName PSObject -Property $Properties | Select-Object *
	
	return $LanguageObject
}
#endregion

#region Variables
[string]$ScriptDirectory = Get-ScriptDirectory
$Module = "ActiveDirectory"
$DomainControllers = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().DomainControllers.Name
$AvailableLanguages = @("English","Nederlands","Deutsch")
$VersionNumber = [System.Windows.Forms.Application]::ProductVersion
$OrgName = "Radboudumc $([char]0x00A9)"
$EmailBlockRule = "2021 Redirect inbound mail from specific SENDERS"
$SubjectBodyBlockRule = "2021 Redirect inbound mail specific keywords in BODY or SUBJECT"
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
	