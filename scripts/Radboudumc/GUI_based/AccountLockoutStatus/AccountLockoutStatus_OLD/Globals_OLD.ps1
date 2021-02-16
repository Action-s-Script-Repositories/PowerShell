#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
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

#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory

$Module = "ActiveDirectory"

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


