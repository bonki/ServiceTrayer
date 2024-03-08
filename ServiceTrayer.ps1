<#
Copyright © 2024 Adrian Frühwirth

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

<#
.SYNOPSIS
	Turns Windows services into a systray icon.

	https://github.com/bonki/ServiceTrayer

.DESCRIPTION
	Copyright (c) 2024 Adrian Frühwirth
	Released under the terms of the MIT license.

	Turns one or more Windows service(s) into a systray icon from which they can be
	started/stopped easily (requires elevated privileges).

	Service status changes can also be watched based on a given interval and if a
	change occurs in the background a notification will be shown to the user.

	Double-clicking the systray icon will start services.msc.

.PARAMETER Help
	Show this help.

.PARAMETER DisplayName
	A regular expression which filters services based on their DisplayName.

.PARAMETER GarbageCollect
	An interval in seconds which specifies how often to run the garbage collector.

	An argument of 0 (zero) disables interval-based garbage collection.

	Default: 600 (10 minutes)

.PARAMETER HideSystray
	Hides the systray icon. This only really makes sense in conjunction with -Watch.

.PARAMETER ListServices
	Whether/how to list services in the context menu. If the program is run with
	elevated privileges and -ListServices is set to 'Manage', the user is allowed
	to start/stop services. If the program is run without elevated privileges
	'Manage' will automatically be degraded to 'Readonly'.
	This parameter is mostly useful in conjunction with -Watch to hide services
	in the context menu if only passive watching is desired.

	Valid arguments: Never | Readonly | Manage
	Default: Manage

.PARAMETER Name
	A regular expression which filters services based on their Name.

.PARAMETER ShowNameAs
	Whether to show service names by their Name or DisplayName.

	Valid arguments: DisplayName | Name
	Default: DisplayName

.PARAMETER ShowNotifications
	When to show notifications.

	Valid arguments: Never | OnError | Always
	Default: OnError

.PARAMETER StartedIcon
	The type of notification icon to show when a service is started.

	Valid arguments: None | Info | Warning | Error
	Default: None

.PARAMETER StoppedIcon
	The type of notification icon to show when a service is stopped.

	Valid arguments: None | Info | Warning | Error
	Default: None

.PARAMETER SystrayIcon
	Path to a file from which to load the icon to use for the systray.
	This uses ExtractAssociatedIcon() and can load an icon from any supported file
	for which Windows can return an associated icon.

	Default: %WINDIR%\System32\services.msc

.PARAMETER Watch
	If given an interval in seconds, shows notifications for status changes for matching
	services happening outside of the application. Note that these notifications are
	always shown irrespective of the value of the -ShowNotifications argument.

	An argument of 0 (zero) has the same effect as not specifying -Watch, i.e. disables
	this feature.

.INPUTS
	None.

.OUTPUTS
	None.

.EXAMPLE
	PS> .\ServiceTrayer.ps1 -DisplayName 'Windows Search'

.EXAMPLE
	PS> .\ServiceTrayer.ps1 -Name '^wuauserv$' -ShowNameAs Name -ShowNotifications Always

.EXAMPLE
	PS> .\ServiceTrayer.ps1 -DisplayName '^Hyper-V' -Watch 20 -StartedIcon Info -StoppedIcon Error -SystrayIcon "%WINDIR%\System32\virtmgmt.msc"
#>

param(
	[string]$DisplayName,
	[int]$GarbageCollect       = 10 * 60,
	[switch]$Help,
	[switch]$HideSystray,
	[string]$ListServices      = "Manage",
	[string]$Name,
	[string]$ShowNameAs        = "DisplayName",
	[string]$ShowNotifications = "OnError",
	[string]$StartedIcon       = "None",
	[string]$StoppedIcon       = "None",
	[string]$SystrayIcon       = "%WINDIR%\System32\services.msc",
	[int]$Watch
)

$About = @{
	Name      = "ServiceTrayer"
	Copyright = "Copyright © 2024 Adrian Frühwirth. All Rights Reserved."
	License   = "Released under the terms of the MIT license."
	URL       = "https://github.com/bonki/ServiceTrayer"
}

enum ListServices {
	Never
	Readonly
	Manage
}

enum ShowNameAs {
	DisplayName
	Name
}

enum ShowNotifications {
	Never
	OnError
	Always
}

class Parameter
{
	[string]$Name
	[object]$Value
	[object]$Transformator
	[object]$Validator

	Parameter($Name, $Value, $Transformator, $Validator) {
	   $this.Name          = $Name
	   $this.Value         = $Value
	   $this.Transformator = $Transformator
	   $this.Validator     = $Validator
	}
}

# early loading so we can access the [System.Windows.Forms.ToolTipIcon] enum for param validation
$null = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$Parameters = @(
	[Parameter]::new("DisplayName",       $DisplayName,       {EmptyRegexTransformator},       [regex]),
	[Parameter]::new("GarbageCollect",    $GarbageCollect,    {SecondsToMsTransformator},      {GarbageCollectValidator}),
	[Parameter]::new("HideSystray",       $HideSystray,       $null,                           $null),
	[Parameter]::new("ListServices",      $ListServices,      {return [ListServices]$_},      [ListServices]),
	[Parameter]::new("Name",              $Name,              {EmptyRegexTransformator},       [regex]),
	[Parameter]::new("ShowNameAs",        $ShowNameAs,        {return [ShowNameAs]$_},        [ShowNameAs]),
	[Parameter]::new("ShowNotifications", $ShowNotifications, {return [ShowNotifications]$_}, [ShowNotifications]),
	[Parameter]::new("StartedIcon",       $StartedIcon,       $null,                           [System.Windows.Forms.ToolTipIcon]),
	[Parameter]::new("StoppedIcon",       $StoppedIcon,       $null,                           [System.Windows.Forms.ToolTipIcon]),
	[Parameter]::new("SystrayIcon",       $SystrayIcon,       $null,                           'filepath'),
	[Parameter]::new("Watch",             $Watch,             {SecondsToMsTransformator},      {WatchValidator})
)

function SecondsToMsTransformator {
	return $_ * 1000
}

function EmptyRegexTransformator {
	if ($_) {
		return $_
	} else {
		return "^$"
	}
}

function GarbageCollectValidator {
	if ($_ -ne 0 -and $_ -lt 60) {
		throw [System.ArgumentException]::new("Value must be 0 | >= 60")
	}
}

function WatchValidator {
	if ($_ -ne 0 -and $_ -lt 5) {
		throw [System.ArgumentException]::new("Value must be 0 | >= 5")
	}
}

# support -Help
if ($Help.IsPresent) {
	Get-Help $MyInvocation.MyCommand.Definition -Detailed
	exit 0
}

# param validation - we're rolling our own so we can do cool stuff(tm)
$Options = @{}
try {
	foreach ($Param in $Parameters) {
		switch ($Param.Validator) {
			'filepath' {
				try {
					if ($Param.Value -ne "") {
						# filepaths auto-expand env variables
						$Param.Value = [System.Environment]::ExpandEnvironmentVariables($Param.Value)
						[System.IO.File]::OpenRead($Param.Value).Close()
					}
				} catch {
					throw [System.ArgumentException]::new("Invalid argument '{0}' for parameter '{1}' of type path: {2}" -f (
						$Param.Value,
						$Param.Name,
						"File not found or not readable"
					)) 
				}
			}
			{$PSItem -eq [regex]} {
				try {
					$null = "" -match $Param.Value
				} catch {
					throw [System.ArgumentException]::new("Invalid argument '{0}' for parameter '{1}' of type regex: {2}" -f (
						$Param.Value,
						$Param.Name,
						$_
					))
				}
			}
		}

		switch ($Param.Validator.BaseType) {
			{$PSItem -eq [enum]} {
				$ValidEnumValues = [Enum]::GetNames($Param.Validator)
				if ($Param.Value -notin $ValidEnumValues) {
					throw [System.ArgumentException]::new("Invalid argument '{0}' for parameter '{1}'. Value must be any of: {2}" -f (
						$Param.Value,
						$Param.Name,
						($ValidEnumValues -join ", ")
					))
				}
			}
		}

		if ($Param.Validator -and $Param.Validator.GetType().Name -eq [ScriptBlock]) {
			try {
				$_ = $Param.Value
				. $Param.Validator
			} catch {
				throw [System.ArgumentException]::new("Invalid argument '{0}' for parameter '{1}': {2}" -f (
					$Param.Value,
					$Param.Name,
					$_
				))
			}
		}

		# transform value and add to Options object
		switch ($Param.Transformator) {
			$null {
				Add-Member -InputObject $Options -MemberType NoteProperty -Name $Param.Name -Value $Param.Value
			}
			default {
				$_ = $Param.Value
				Add-Member -InputObject $Options -MemberType NoteProperty -Name $Param.Name -Value (. $Param.Transformator)
			}
		}
	}

	# we check this last so parameters failing validation will error out first
	# we don't check their respective $Options value because they are already transformed and checking for ^$
	# might break in the future if the defaults should change
	if ($DisplayName -eq "" -and $Name -eq "") {
		throw [System.ArgumentException]::new(
			"Missing parameter: Must specify at least one of -DisplayName or -Name." +
			"`nIf you really wish to match all services you may force this using e.g. -Name '.' or similar."
		)
	}
} catch {
	Write-Host ("Error: {0}`nRun '{1} -Help' for more info." -f ($_, $MyInvocation.MyCommand.Definition))
	exit 1
}

# @fixme
# this hides the console window - takes a long time, why? is loading of the initial PS env really that slow?
$WindowCode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
$AsyncWindow = Add-Type -MemberDefinition $WindowCode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru
$null = $AsyncWindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)

function NewAboutDialog {
	$Dialog = [System.Windows.Forms.Form]@{
		Text            = "About {0}" -f $About.Name
		Icon            = $Sync.Tray.Icon
		Font            = "Segoe UI,9"
		Size            = New-Object System.Drawing.Size(400, 280)
		FormBorderStyle = "FixedSingle"
		MaximizeBox     = $false
		MinimizeBox     = $false
	}

	# TableLayoutPanel to center elements easily. centering requires elements to explicitly set their Anchor to 'None'
	$Layout = [System.Windows.Forms.TableLayoutPanel]@{
		Dock            = "Fill"
		ColumnCount     = 1
		RowCount        = 7
		#CellBorderSTyle = "single" # for debugging purposes
	}

	$null = $Layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Autosize)))
	$null = $Layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 13)))
	$null = $Layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25)))
	$null = $Layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 15)))
	$null = $Layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$null = $Layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
	$null = $Layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
	$null = $Layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 5)))

	$Title = @{
		Text = $About.Name
		Font = "Segoe UI,24,style=Bold"
	}
	$TitleLabel = [System.Windows.Forms.Label]@{
		Text      = $Title.Text
		Font      = $Title.Font
		AutoSize  = $true
		TextAlign = "MiddleCenter"
		Anchor    = "None"
	}

	$URLLabel = [System.Windows.Forms.LinkLabel]@{
		Text      = $About.URL
		AutoSize  = $true
		TextAlign = "MiddleCenter"
		Anchor    = "None"
	}
	$URLLabel.Add_Click({
		Start-Process $URLLabel.Text
	}.GetNewClosure())

	$CopyrightLabel = [System.Windows.Forms.Label]@{
		Text      = $About.Copyright
		AutoSize  = $true
		TextAlign = "MiddleCenter"
		Anchor    = "None"
	}

	$LicenseLabel = [System.Windows.Forms.Label]@{
		Text      = $About.License
		AutoSize  = $true
		TextAlign = "MiddleCenter"
		Anchor    = "None"
	}

	$CloseButton = [System.Windows.Forms.Button]@{
		Text      = "Close"
		FlatStyle = "Flat"
		Anchor    = "None"
	}
	$CloseButton.Add_Click({
		$Dialog.Close()
	}.GetNewClosure())

	$i = 1
	foreach ($item in @(
		$TitleLabel,
		$URLLabel,
		$CopyrightLabel,
		$LicenseLabel,
		$CloseButton
	)) {
		$Layout.Controls.Add($item, 0, $i++)
	}

	$EasterEggQuotes = @(
		"I once saved the entire planet`nwith a mop and some bubblegum.",
		"Look behind you, a three-headed monkey!",
		"How much wood could a woodchuck chuck`nif a woodchuck could chuck wood?",
		"I am selling these fine leather jackets.",
		"Wax fruit? My favorite!",
		"We're not lost.`nWe're locationally challenged.",
		"Reality is what you make it,`nso make it funny.",
		"I asked my cat for life advice. It suggested`nnapping. Wise counsel, but I’m`nmore of a caffeine-fueled adventurer."
	)

	$UpdateEasterEgg = {
		$TitleLabel.Tag  = ($TitleLabel.Tag + 1) % $EasterEggQuotes.Length
		$TitleLabel.Font = "Segoe UI,12,style=Bold,Italic"
		$TitleLabel.Text = '"' + $EasterEggQuotes[$TitleLabel.Tag] + '"'
	}.GetNewClosure()

	$TitleLabel.Add_MouseDown($UpdateEasterEgg)
	$Layout.Add_MouseDown($UpdateEasterEgg)

	$Dialog.Add_Closing({
		$TitleLabel.Tag  = 0
		$TitleLabel.Text = $Title.Text
		$TitleLabel.Font = $Title.Font
	}.GetNewClosure())

	$Dialog.Controls.Add($Layout)

	return $Dialog
}

function IsAdmin {
	return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
}

function GetMatchingServices {
	Get-Service | Where { ($_.DisplayName -match $Options.DisplayName) -or ($_.Name -match $Options.Name) }
}

function MenuClickHandler {
	# when starting/stopping a service via the context menu we kick off an asynchronous runspace job so that the GUI doesn't block
	$global:Session = [PowerShell]::Create().AddScript({
		param($Sync)

		$Options = $Sync.Options
		$Tray    = $Sync.Tray
		$Item    = $Sync.Item
		$Service = Get-Service -Name $Item.Name

		# we add the current service to a temporary ignore list so that the watch timer won't act
		# on status changes the user initiated, otherwise they would get an additional (delayed) notification.
		# the timer is only meant to watch for changes happening *outside* of our application.
		if ($Options.Watch -gt 0 -and $Service.Name -notin $Sync.TimerIgnoreStatusChange) {
			$Sync.TimerIgnoreStatusChange.Add($Service.Name)
		}
		
		$ServiceName = switch ($Options.ShowNameAs) {
			DisplayName { $Service.DisplayName }
			Name        { $Service.Name        }
		}

		try {
			switch ($Service.Status) {
				"Running" {
					$Service | Stop-Service -ErrorAction Stop
					$Service.WaitForStatus("Stopped", 0)
					if ($Options.ShowNotifications -eq "Always") {
						$Tray.ShowBalloonTip(5000, $Tray.Text, "Service '{0}' stopped" -f $ServiceName, [system.windows.forms.ToolTipIcon]$Options.StoppedIcon)
					}
				}
				"Stopped" {
					$Service | Start-Service -ErrorAction Stop
					$Service.WaitForStatus("Running", 0)
					if ($Options.ShowNotifications -eq "Always") {
						$Tray.ShowBalloonTip(5000, $Tray.Text, "Service '{0}' started" -f $ServiceName, [system.windows.forms.ToolTipIcon]$Options.StartedIcon)
					}
				}
			}
		} catch {
			if ($Options.ShowNotifications -ge "OnError") {
				$Tray.ShowBalloonTip(5000, $Tray.Text, $_, [system.windows.forms.ToolTipIcon]"Error")
			}
		}
	}).AddArgument($Sync).BeginInvoke()
}

function UpdateContextMenu {
	$Options  = $Sync.Options
	$Services = GetMatchingServices

	# redrawing the menu is expensive and potentially leaky - don't do it unless something has changed
	if (($Script:MenuPreviousServices -ne $null) -and ((Compare-Object $Script:MenuPreviousServices $Services -Property DisplayName, Name, Status) -eq $null)) {
		return
	}
	$Script:MenuPreviousServices = $Services

	$Tray = $Sync.Tray
	$ContextMenu = $Tray.ContextMenu
	$ContextMenu.MenuItems.Clear()

	$IsAdmin = IsAdmin

	if ($Options.ListServices -ge "Readonly") {
		$Services | ForEach-Object {
			$Service = $_
			$ServiceName = switch ($Sync.Options.ShowNameAs) {
				DisplayName { $Service.DisplayName }
				Name        { $Service.Name        }
			}
			$AllowManage = $IsAdmin -and $Options.ListServices -eq "Manage" # starting/stopping services requires elevated privileges
			$ServiceItem = [System.Windows.Forms.MenuItem]@{
				Text    = $ServiceName
				Name    = $Service.Name
				Checked = $Service.Status -in @("Running")
				Enabled = $AllowManage -and $Service.Status -in @("Running", "Stopped")
			}

			$ServiceItem.Add_Click({
				$Sync.Item = $this
				MenuClickHandler
			})
			$ContextMenu.MenuItems.AddRange($ServiceItem)
		}

		if ($Services.Length -ne 0) {
			$ContextMenu.MenuItems.AddRange("-") # this is win32 magic for drawing a spacer
		}
	}

	$OpenServicesItem = New-Object System.Windows.Forms.MenuItem
	$OpenServicesItem.Text = "Open Services..."
	$OpenServicesItem.Add_Click({
		Start-Process services.msc
	})
	$ContextMenu.MenuItems.AddRange($OpenServicesItem)
	$ContextMenu.MenuItems.AddRange("-")

	$AboutItem = New-Object System.Windows.Forms.MenuItem
	$AboutItem.Text = "About {0}..." -f $About.Name
	$AboutItem.Add_Click({
		if (!$AboutDialog.Visible) {
			$Screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
			$CenterLeft = ($Screen.Width  - $AboutDialog.Width)  / 2
			$CenterTop  = ($Screen.Height - $AboutDialog.Height) / 2
			$AboutDialog.StartPosition = "Manual"
			$AboutDialog.Location = New-Object System.Drawing.Point($CenterLeft, $CenterTop)

			$AboutDialog.ShowDialog()
		}
	})
	$ContextMenu.MenuItems.AddRange($AboutItem)

	$ExitItem = New-Object System.Windows.Forms.MenuItem
	$ExitItem.Text = "Exit"
	$ExitItem.Add_Click({
		[System.Windows.Forms.Application]::Exit()
	})
	$ContextMenu.MenuItems.AddRange($ExitItem)

	# run garbage collector on redraw if GC timer is disabled
	if ($Options.GarbageCollect -eq 0) {
		[System.GC]::Collect()
	}
}

[System.GC]::Collect()

$null = [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
try {
	$TrayIcon = [System.Drawing.Icon]::ExtractAssociatedIcon($Options.SystrayIcon)
} catch [System.IO.FileNotFoundException] {
	Write-Host ("Failed to load icon from file '{0}': {1}" -f ($Options.SystrayIcon, "File not found"))
	exit 1
} catch {
	Write-Host ("Failed to load icon from file '{0}': {1}" -f ($Options.SystrayIcon, $_))
	exit 1
}

# synchronized hashtable for passing stuff into the runspace jobs
$Sync = [Hashtable]::Synchronized(@{})
$Sync.Options = $Options
$Sync.Tray = [System.Windows.Forms.NotifyIcon]@{
	Text    = $About.Name
	Icon    = $TrayIcon
	Visible = !$Options.HideSystray
}

$AboutDialog = NewAboutDialog

# @todo
# this, if requested, enables a timer which each $Options.Watch seconds checks if any of our
# services of interest has transitioned to another state so that we can notify the user.
# ideally we should use something like NotifyServiceStatusChangeW()/SubscribeServiceChangeNotifications()
# to subscribe to service changes instead because polling is meh for obvious reasons.
if ($Options.Watch -gt 0) {
	$Sync.TimerIgnoreStatusChange = New-Object System.Collections.ArrayList
	$Sync.Timer = [System.Windows.Forms.Timer]@{
		Interval = $Options.Watch
	}
	$Sync.TimerPreviousServices = GetMatchingServices
	$Sync.Timer.Add_Tick({
		if ($Sync.TimerPreviousServices -eq $null) {
			return
		}

		$Services = GetMatchingServices

		foreach ($Service in Compare-Object $Sync.TimerPreviousServices $Services -Property DisplayName, Name, Status -PassThru ` | Where { `
			     $_.SideIndicator -eq "=>" `
			-and $_.Status -in @("Running", "Stopped") `
		}) {
			# see above - when initiating a status change from within the application we don't want the timer
			# to fire notifications for those. if such a change is detected, ignore it and remove the service
			# from the ignore list
			if ($Service.Name -in $Sync.TimerIgnoreStatusChange) {
				$Sync.TimerIgnoreStatusChange.Remove($Service.Name)
				continue
			}

			$ServiceName = switch ($Sync.Options.ShowNameAs) {
				DisplayName { $Service.DisplayName }
				Name        { $Service.Name        }
			}
			Switch ($Service.Status) {
				"Running" {
					$Sync.Tray.ShowBalloonTip(5000, $Sync.Tray.Text, "Service '{0}' started" -f $ServiceName, [system.windows.forms.ToolTipIcon]$Sync.Options.StartedIcon)
				}
				"Stopped" {
					$Sync.Tray.ShowBalloonTip(5000, $Sync.Tray.Text, "Service '{0}' stopped" -f $ServiceName, [system.windows.forms.ToolTipIcon]$Sync.Options.StoppedIcon)
				}
			}
		}

		$Sync.TimerPreviousServices = $Services
	})

	$Sync.Timer.Start()
}

# yuck
if ($Options.GarbageCollect -gt 0) {
	$GCTimer = [System.Windows.Forms.Timer]@{
		Interval = $Options.GarbageCollect
	}
	$GCTimer.Add_Tick({
		[System.GC]::Collect()
	})
	$GCTimer.Start()
}

# doesn't seem to be needed?
#$Runspace = [RunspaceFactory]::CreateRunspace()
#$Runspace.ApartmentState = [Threading.ApartmentState]::STA
#$Runspace.Open()

$Sync.Tray.ContextMenu = New-Object System.Windows.Forms.ContextMenu
UpdateContextMenu
# @xxx
# when showing notifications this doesn't seem to fire so the user has to click twice for
# the menu to update, without notifications this works just fine - why?
$Sync.Tray.Add_MouseDown({
	UpdateContextMenu
})

$Sync.Tray.Add_DoubleClick({
	if ($_.Button -ne [Windows.Forms.MouseButtons]::Left) {
		return
	}

	Start-Process services.msc
})

# let CTRL-C not be a Signal so we can't crash the GUI thread when run from the command line
[console]::TreatControlCAsInput = $true
$AppContext = New-Object System.Windows.Forms.ApplicationContext
[void][System.Windows.Forms.Application]::Run($AppContext)
$Host.UI.RawUI.FlushInputBuffer()

# cleanup - we should really shove this into some kind of exit handler (is there?) so we can clean up when
# exiting ungracefully?
# there is [System.Windows.Forms.Application]::Add_ApplicationExit() but there is no benefit in using that
# because it won't be called when exiting ungracefully
# there is also [System.Windows.Forms.Application]::Add_ThreadException() but I'm not sure how to recover
# from within such a handler. it does silence the GUI dialog showing the stack trace but the message loop
# has already crashed at that point and the UX is worse compared to simply disabling CTRL-C
if ($Sync.Tray.Icon) {
	$Sync.Tray.Icon.Dispose()
}
if ($Sync.Tray.ContextMenu) {
	$Sync.Tray.ContextMenu.Dispose()
}
if ($Sync.Tray) {
	$Sync.Tray.Dispose()
}
if ($Sync.Timer) {
	$Sync.Timer.Stop()
	$Sync.Timer.Dispose()
}
if ($Sync.GCTimer) {
	$Sync.GCTimer.Stop()
	$Sync.GCTimer.Dispose()
}
