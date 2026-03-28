param(
    [switch]$SmokeTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http

Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class WallpaperNative
{
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SystemParametersInfo(
        int uAction,
        int uParam,
        string lpvParam,
        int fuWinIni
    );

    [DllImport("gdi32.dll", SetLastError = true)]
    public static extern IntPtr CreateRoundRectRgn(
        int nLeftRect,
        int nTopRect,
        int nRightRect,
        int nBottomRect,
        int nWidthEllipse,
        int nHeightEllipse
    );

    [DllImport("gdi32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DeleteObject(IntPtr hObject);
}
"@

$script:ConfigPath = Join-Path $PSScriptRoot 'config.json'
$script:CacheDirectory = Join-Path (Join-Path $env:LOCALAPPDATA 'WallpaperPuller') 'Cache'
$script:State = @{
    SearchUrl = $null
    ApiUrl = $null
    WallpaperStyle = 'Fill'
    CurrentFilePath = $null
}

$script:CanvasColor = [System.Drawing.ColorTranslator]::FromHtml('#E6EFEC')
$script:SurfaceColor = [System.Drawing.ColorTranslator]::FromHtml('#F8FBFA')
$script:SoftBorderColor = [System.Drawing.ColorTranslator]::FromHtml('#CCDBD6')
$script:SidebarStartColor = [System.Drawing.ColorTranslator]::FromHtml('#123E38')
$script:SidebarEndColor = [System.Drawing.ColorTranslator]::FromHtml('#245951')
$script:SidebarCardColor = [System.Drawing.ColorTranslator]::FromHtml('#194B43')
$script:AccentBlueColor = [System.Drawing.ColorTranslator]::FromHtml('#8FD9FF')
$script:AccentGreenColor = [System.Drawing.ColorTranslator]::FromHtml('#B7ED79')
$script:AccentBlueHoverColor = [System.Drawing.ColorTranslator]::FromHtml('#ACE4FF')
$script:AccentGreenHoverColor = [System.Drawing.ColorTranslator]::FromHtml('#CCF39C')
$script:AccentBluePressedColor = [System.Drawing.ColorTranslator]::FromHtml('#70C7F3')
$script:AccentGreenPressedColor = [System.Drawing.ColorTranslator]::FromHtml('#A0D764')
$script:TextPrimaryColor = [System.Drawing.ColorTranslator]::FromHtml('#142229')
$script:TextSecondaryColor = [System.Drawing.ColorTranslator]::FromHtml('#60707A')
$script:TextOnDarkColor = [System.Drawing.ColorTranslator]::FromHtml('#EEF6F3')
$script:TextMutedOnDarkColor = [System.Drawing.ColorTranslator]::FromHtml('#BED1CB')
$script:NeutralColor = [System.Drawing.ColorTranslator]::FromHtml('#75A5B4')
$script:SuccessColor = [System.Drawing.ColorTranslator]::FromHtml('#57B874')
$script:ErrorColor = [System.Drawing.ColorTranslator]::FromHtml('#D35B5B')

function Get-DefaultConfig {
    return [pscustomobject]@{
        WallhavenSearchUrl = 'https://wallhaven.cc/search?q=canada%20nature&categories=100&purity=100&atleast=2560x1440&sorting=date_added&order=desc&page=2'
        WallpaperStyle = 'Fill'
    }
}

function Initialize-ConfigFile {
    if (-not (Test-Path -LiteralPath $script:ConfigPath)) {
        Get-DefaultConfig | ConvertTo-Json | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
    }
}

function Get-AppConfig {
    Initialize-ConfigFile
    $config = Get-Content -LiteralPath $script:ConfigPath -Raw | ConvertFrom-Json

    if (-not $config.WallhavenSearchUrl) {
        throw 'config.json is missing WallhavenSearchUrl.'
    }

    if (-not $config.WallpaperStyle) {
        $config | Add-Member -MemberType NoteProperty -Name WallpaperStyle -Value 'Fill'
    }

    return $config
}

function Convert-WallhavenSearchUrlToApiUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchUrl
    )

    $uri = [System.Uri]::new($SearchUrl)
    if ($uri.Host -notmatch '(^|\.)wallhaven\.cc$') {
        throw 'The configured URL must point to wallhaven.cc.'
    }

    if ($uri.AbsolutePath -eq '/api/v1/search') {
        return $uri.AbsoluteUri
    }

    if ($uri.AbsolutePath -ne '/search') {
        throw 'Expected a Wallhaven search URL ending in /search.'
    }

    $builder = [System.UriBuilder]::new($uri)
    $builder.Path = '/api/v1/search'
    return $builder.Uri.AbsoluteUri
}

function Initialize-CacheDirectory {
    New-Item -ItemType Directory -Path $script:CacheDirectory -Force | Out-Null
}

function New-HttpClient {
    $client = [System.Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromSeconds(30)
    $client.DefaultRequestHeaders.UserAgent.ParseAdd('WallpaperPuller/1.0')
    return $client
}

function Get-SafePropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        $Object,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $Object) {
        return $null
    }

    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}
function Get-DisplayValue {
    param(
        $Value,
        [string]$Fallback = 'Not available'
    )

    if ($null -eq $Value) { return $Fallback }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $Fallback }
    return $text
}


function Set-RoundedRegion {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Control]$Control,
        [int]$Radius = 18
    )

    if ($Control.Width -le 0 -or $Control.Height -le 0) {
        return
    }

    $handle = [WallpaperNative]::CreateRoundRectRgn(0, 0, $Control.Width + 1, $Control.Height + 1, $Radius * 2, $Radius * 2)
    if ($handle -eq [IntPtr]::Zero) {
        return
    }

    try {
        $Control.Region = [System.Drawing.Region]::FromHrgn($handle)
    }
    finally {
        [WallpaperNative]::DeleteObject($handle) | Out-Null
    }
}

function Enable-RoundedCorners {
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Control]$Control,
        [int]$Radius = 18
    )

    Set-RoundedRegion -Control $Control -Radius $Radius
}


function Update-ActionButtons {
    if ($NewPictureButton.Enabled) {
        $NewPictureButton.BackColor = $script:AccentBlueColor
        $NewPictureButton.ForeColor = $script:TextPrimaryColor
        $NewPictureButton.FlatAppearance.BorderColor = $script:AccentBluePressedColor
        $NewPictureButton.FlatAppearance.MouseOverBackColor = $script:AccentBlueHoverColor
        $NewPictureButton.FlatAppearance.MouseDownBackColor = $script:AccentBluePressedColor
    }
    else {
        $NewPictureButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#D7E2DE')
        $NewPictureButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#86938E')
        $NewPictureButton.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#C1CEC8')
    }

    if ($SetWallpaperButton.Enabled) {
        $SetWallpaperButton.BackColor = $script:AccentGreenColor
        $SetWallpaperButton.ForeColor = $script:TextPrimaryColor
        $SetWallpaperButton.FlatAppearance.BorderColor = $script:AccentGreenPressedColor
        $SetWallpaperButton.FlatAppearance.MouseOverBackColor = $script:AccentGreenHoverColor
        $SetWallpaperButton.FlatAppearance.MouseDownBackColor = $script:AccentGreenPressedColor
    }
    else {
        $SetWallpaperButton.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#D7E2DE')
        $SetWallpaperButton.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#86938E')
        $SetWallpaperButton.FlatAppearance.BorderColor = [System.Drawing.ColorTranslator]::FromHtml('#C1CEC8')
    }
}

function Reset-WallpaperDetails {
    $MetaResolutionLabel.Text = 'Resolution: waiting'
    $MetaIdLabel.Text = 'Wallpaper ID: waiting'
    $MetaTypeLabel.Text = 'Format: waiting'
    $MetaStyleLabel.Text = ('Mode: ' + $script:State.WallpaperStyle)
}

function Update-WallpaperDetails {
    param(
        [Parameter(Mandatory = $true)]
        $Wallpaper
    )

    $identifier = Get-DisplayValue -Value (Get-SafePropertyValue -Object $Wallpaper -Name 'id') -Fallback 'unknown'
    $width = Get-DisplayValue -Value (Get-SafePropertyValue -Object $Wallpaper -Name 'dimension_x') -Fallback '?'
    $height = Get-DisplayValue -Value (Get-SafePropertyValue -Object $Wallpaper -Name 'dimension_y') -Fallback '?'
    $fileType = Get-DisplayValue -Value (Get-SafePropertyValue -Object $Wallpaper -Name 'file_type') -Fallback 'unknown'

    $MetaResolutionLabel.Text = ('Resolution: {0}x{1}' -f $width, $height)
    $MetaIdLabel.Text = ('Wallpaper ID: {0}' -f $identifier)
    $MetaTypeLabel.Text = ('Format: {0}' -f $fileType.Replace('image/', '').ToUpperInvariant())
    $MetaStyleLabel.Text = ('Mode: ' + $script:State.WallpaperStyle)
}
function Set-Status {
    param(
        [string]$Message,
        [System.Drawing.Color]$Color
    )

    $StatusLabel.Text = $Message
    $StatusLabel.ForeColor = $script:TextOnDarkColor
    $StatusAccentBar.BackColor = $Color
    if (Get-Variable -Name PreviewMetaInfoLabel -Scope Script -ErrorAction SilentlyContinue) {
        $PreviewMetaInfoLabel.Text = $Message
    }
    $StatusLabel.Refresh()
}

function Clear-Preview {
    if ($PictureBox.Image) {
        $oldImage = $PictureBox.Image
        $PictureBox.Image = $null
        $oldImage.Dispose()
    }

    $PlaceholderLabel.Visible = $true
}

function Show-PreviewImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($PictureBox.Image) {
        $oldImage = $PictureBox.Image
        $PictureBox.Image = $null
        $oldImage.Dispose()
    }

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $image = [System.Drawing.Image]::FromStream($stream)
        try {
            $PictureBox.Image = [System.Drawing.Bitmap]::new($image)
        }
        finally {
            $image.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }

    $PlaceholderLabel.Visible = $false
}

function Set-BusyState {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$IsBusy
    )

    $NewPictureButton.Enabled = -not $IsBusy
    $SetWallpaperButton.Enabled = (-not $IsBusy) -and [bool]$script:State.CurrentFilePath
    $MainForm.UseWaitCursor = $IsBusy
    $MainForm.Cursor = if ($IsBusy) { [System.Windows.Forms.Cursors]::WaitCursor } else { [System.Windows.Forms.Cursors]::Default }
    Update-ActionButtons
    $MainForm.Refresh()
}

function Get-RandomWallpaper {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ApiUrl
    )

    $client = New-HttpClient
    try {
        $response = $client.GetAsync($ApiUrl).GetAwaiter().GetResult()
        $null = $response.EnsureSuccessStatusCode()
        $payload = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult() | ConvertFrom-Json
    }
    finally {
        $client.Dispose()
    }

    $items = @((Get-SafePropertyValue -Object $payload -Name 'data'))
    if ($items.Count -eq 0) {
        throw 'Wallhaven returned no results for the configured search.'
    }

    $downloadableItems = @(
        $items | Where-Object {
            $candidatePath = Get-SafePropertyValue -Object $_ -Name 'path'
            -not [string]::IsNullOrWhiteSpace([string]$candidatePath)
        }
    )

    if ($downloadableItems.Count -eq 0) {
        throw 'Wallhaven returned results, but none included a downloadable image path.'
    }

    return ($downloadableItems | Get-Random)
}

function Save-WallpaperToCache {
    param(
        [Parameter(Mandatory = $true)]
        $Wallpaper
    )

    $imagePath = [string](Get-SafePropertyValue -Object $Wallpaper -Name 'path')
    if ([string]::IsNullOrWhiteSpace($imagePath)) {
        $typeName = if ($null -eq $Wallpaper) { 'null' } else { $Wallpaper.GetType().FullName }
        $availableProperties = if ($null -eq $Wallpaper) { '' } else { ($Wallpaper.PSObject.Properties.Name -join ', ') }
        throw ('Wallhaven item did not include a downloadable image path. Type: {0}. Properties: {1}' -f $typeName, $availableProperties)
    }

    Initialize-CacheDirectory

    $wallpaperUri = [System.Uri]::new($imagePath)
    $extension = [System.IO.Path]::GetExtension($wallpaperUri.AbsolutePath)
    if ([string]::IsNullOrWhiteSpace($extension)) {
        $extension = '.jpg'
    }

    $fileId = [string](Get-SafePropertyValue -Object $Wallpaper -Name 'id')
    if ([string]::IsNullOrWhiteSpace($fileId)) {
        $fileId = [guid]::NewGuid().ToString('N')
    }

    $destination = Join-Path $script:CacheDirectory ($fileId + $extension)

    if (Test-Path -LiteralPath $destination) {
        return $destination
    }

    $temporaryPath = $destination + '.download'
    if (Test-Path -LiteralPath $temporaryPath) {
        Remove-Item -LiteralPath $temporaryPath -Force
    }

    $client = New-HttpClient
    try {
        $response = $client.GetAsync($wallpaperUri, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
        $null = $response.EnsureSuccessStatusCode()

        $responseStream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        try {
            $targetStream = [System.IO.File]::Open($temporaryPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            try {
                $responseStream.CopyTo($targetStream)
            }
            finally {
                $targetStream.Dispose()
            }
        }
        finally {
            $responseStream.Dispose()
        }

        Move-Item -LiteralPath $temporaryPath -Destination $destination -Force
    }
    finally {
        $client.Dispose()
    }

    return $destination
}

function Set-WallpaperStyle {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Fill', 'Fit', 'Stretch', 'Center', 'Span', 'Tile')]
        [string]$Style
    )

    $desktopKey = 'HKCU:\Control Panel\Desktop'
    switch ($Style) {
        'Fill' { Set-ItemProperty -Path $desktopKey -Name WallpaperStyle -Value '10'; Set-ItemProperty -Path $desktopKey -Name TileWallpaper -Value '0' }
        'Fit' { Set-ItemProperty -Path $desktopKey -Name WallpaperStyle -Value '6'; Set-ItemProperty -Path $desktopKey -Name TileWallpaper -Value '0' }
        'Stretch' { Set-ItemProperty -Path $desktopKey -Name WallpaperStyle -Value '2'; Set-ItemProperty -Path $desktopKey -Name TileWallpaper -Value '0' }
        'Center' { Set-ItemProperty -Path $desktopKey -Name WallpaperStyle -Value '0'; Set-ItemProperty -Path $desktopKey -Name TileWallpaper -Value '0' }
        'Span' { Set-ItemProperty -Path $desktopKey -Name WallpaperStyle -Value '22'; Set-ItemProperty -Path $desktopKey -Name TileWallpaper -Value '0' }
        'Tile' { Set-ItemProperty -Path $desktopKey -Name WallpaperStyle -Value '0'; Set-ItemProperty -Path $desktopKey -Name TileWallpaper -Value '1' }
    }
}

function Set-DesktopWallpaper {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $absolutePath = [System.IO.Path]::GetFullPath($Path)
    $desktopKey = 'HKCU:\Control Panel\Desktop'

    Set-WallpaperStyle -Style $script:State.WallpaperStyle
    Set-ItemProperty -Path $desktopKey -Name Wallpaper -Value $absolutePath

    $flags = 0x01 -bor 0x02
    $spiSucceeded = [WallpaperNative]::SystemParametersInfo(20, 0, $absolutePath, $flags)
    $refreshProcess = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\rundll32.exe') -ArgumentList 'user32.dll,UpdatePerUserSystemParameters' -WindowStyle Hidden -PassThru
    $refreshProcess.WaitForExit()

    if (-not $spiSucceeded) {
        $errorCode = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
        throw ('Windows did not confirm the wallpaper change (SystemParametersInfo error {0}).' -f $errorCode)
    }
}

function Format-WallpaperSummary {
    param(
        [Parameter(Mandatory = $true)]
        $Wallpaper
    )

    $identifier = [string](Get-SafePropertyValue -Object $Wallpaper -Name 'id')
    if ([string]::IsNullOrWhiteSpace($identifier)) { $identifier = 'unknown' }
    $width = [string](Get-SafePropertyValue -Object $Wallpaper -Name 'dimension_x')
    if ([string]::IsNullOrWhiteSpace($width)) { $width = '?' }
    $height = [string](Get-SafePropertyValue -Object $Wallpaper -Name 'dimension_y')
    if ([string]::IsNullOrWhiteSpace($height)) { $height = '?' }
    return ('Ready: {0} - {1}x{2}' -f $identifier, $width, $height)
}

function Load-NewWallpaper {
    try {
        Set-BusyState -IsBusy $true
        Set-Status -Message 'Pulling a fresh image from Wallhaven...' -Color $script:NeutralColor
        $wallpaper = Get-RandomWallpaper -ApiUrl $script:State.ApiUrl
        Set-Status -Message 'Downloading the selected image...' -Color $script:NeutralColor
        $localPath = Save-WallpaperToCache -Wallpaper $wallpaper
        Show-PreviewImage -Path $localPath
        $script:State.CurrentFilePath = $localPath
        Update-WallpaperDetails -Wallpaper $wallpaper
        Set-Status -Message (Format-WallpaperSummary -Wallpaper $wallpaper) -Color $script:SuccessColor
    }
    catch {
        Clear-Preview
        $script:State.CurrentFilePath = $null
        Reset-WallpaperDetails
        Set-Status -Message $_.Exception.Message -Color $script:ErrorColor
    }
    finally {
        Set-BusyState -IsBusy $false
    }
}

[System.Windows.Forms.Application]::EnableVisualStyles()
$MainForm = New-Object System.Windows.Forms.Form
$MainForm.Text = 'Wallpaper Puller'
$MainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$MainForm.ClientSize = New-Object System.Drawing.Size(1240, 760)
$MainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$MainForm.MaximizeBox = $false
$MainForm.MinimizeBox = $true
$MainForm.BackColor = $script:CanvasColor
$MainForm.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$RootPanel = New-Object System.Windows.Forms.Panel
$RootPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$RootPanel.Padding = New-Object System.Windows.Forms.Padding(18)
$RootPanel.BackColor = $script:CanvasColor
$MainForm.Controls.Add($RootPanel)

$SurfacePanel = New-Object System.Windows.Forms.Panel
$SurfacePanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$SurfacePanel.BackColor = $script:SurfaceColor
$RootPanel.Controls.Add($SurfacePanel)

$LeftPanel = New-Object System.Windows.Forms.Panel
$LeftPanel.Dock = [System.Windows.Forms.DockStyle]::Left
$LeftPanel.Width = 300
$LeftPanel.BackColor = $script:SidebarStartColor
$LeftPanel.Add_Paint({
    param($sender, $eventArgs)
    $rect = $sender.ClientRectangle
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $script:SidebarStartColor, $script:SidebarEndColor, 90)
    try { $eventArgs.Graphics.FillRectangle($brush, $rect) } finally { $brush.Dispose() }
})
$SurfacePanel.Controls.Add($LeftPanel)

$PreviewHost = New-Object System.Windows.Forms.Panel
$PreviewHost.Dock = [System.Windows.Forms.DockStyle]::Fill
$PreviewHost.Padding = New-Object System.Windows.Forms.Padding(30, 24, 30, 24)
$PreviewHost.BackColor = $script:SurfaceColor
$SurfacePanel.Controls.Add($PreviewHost)

$SidebarAccent = New-Object System.Windows.Forms.Panel
$SidebarAccent.Size = New-Object System.Drawing.Size(82, 6)
$SidebarAccent.Location = New-Object System.Drawing.Point(28, 28)
$SidebarAccent.BackColor = $script:AccentBlueColor
Enable-RoundedCorners -Control $SidebarAccent -Radius 3
$LeftPanel.Controls.Add($SidebarAccent)

$SidebarTagLabel = New-Object System.Windows.Forms.Label
$SidebarTagLabel.AutoSize = $true
$SidebarTagLabel.Location = New-Object System.Drawing.Point(28, 50)
$SidebarTagLabel.Text = 'CANADA NATURE FEED'
$SidebarTagLabel.ForeColor = $script:AccentBlueColor
$SidebarTagLabel.BackColor = [System.Drawing.Color]::Transparent
$SidebarTagLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$LeftPanel.Controls.Add($SidebarTagLabel)

$SidebarTitleLabel = New-Object System.Windows.Forms.Label
$SidebarTitleLabel.AutoSize = $true
$SidebarTitleLabel.Location = New-Object System.Drawing.Point(28, 78)
$SidebarTitleLabel.Text = 'Wallpaper Puller'
$SidebarTitleLabel.ForeColor = $script:TextOnDarkColor
$SidebarTitleLabel.BackColor = [System.Drawing.Color]::Transparent
$SidebarTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 24)
$LeftPanel.Controls.Add($SidebarTitleLabel)

$SidebarSubtitleLabel = New-Object System.Windows.Forms.Label
$SidebarSubtitleLabel.Location = New-Object System.Drawing.Point(28, 122)
$SidebarSubtitleLabel.Size = New-Object System.Drawing.Size(242, 74)
$SidebarSubtitleLabel.Text = 'Pull a random Canadian nature wallpaper, preview it, and drop it onto your desktop in one click.'
$SidebarSubtitleLabel.ForeColor = $script:TextMutedOnDarkColor
$SidebarSubtitleLabel.BackColor = [System.Drawing.Color]::Transparent
$SidebarSubtitleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10.5)
$LeftPanel.Controls.Add($SidebarSubtitleLabel)

$NewPictureButton = New-Object System.Windows.Forms.Button
$NewPictureButton.Text = 'New Picture'
$NewPictureButton.Size = New-Object System.Drawing.Size(244, 60)
$NewPictureButton.Location = New-Object System.Drawing.Point(28, 232)
$NewPictureButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$NewPictureButton.FlatAppearance.BorderSize = 1
$NewPictureButton.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
$NewPictureButton.Cursor = [System.Windows.Forms.Cursors]::Hand
Enable-RoundedCorners -Control $NewPictureButton -Radius 18
$LeftPanel.Controls.Add($NewPictureButton)

$ActionHintLabel = New-Object System.Windows.Forms.Label
$ActionHintLabel.Location = New-Object System.Drawing.Point(32, 300)
$ActionHintLabel.Size = New-Object System.Drawing.Size(238, 38)
$ActionHintLabel.Text = 'Keep rolling until you land on a wallpaper worth keeping.'
$ActionHintLabel.ForeColor = $script:TextMutedOnDarkColor
$ActionHintLabel.BackColor = [System.Drawing.Color]::Transparent
$ActionHintLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
$LeftPanel.Controls.Add($ActionHintLabel)

$SetWallpaperButton = New-Object System.Windows.Forms.Button
$SetWallpaperButton.Text = 'Set as Wallpaper'
$SetWallpaperButton.Size = New-Object System.Drawing.Size(244, 60)
$SetWallpaperButton.Location = New-Object System.Drawing.Point(28, 364)
$SetWallpaperButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$SetWallpaperButton.FlatAppearance.BorderSize = 1
$SetWallpaperButton.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 13)
$SetWallpaperButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$SetWallpaperButton.Enabled = $false
Enable-RoundedCorners -Control $SetWallpaperButton -Radius 18
$LeftPanel.Controls.Add($SetWallpaperButton)
$ApplyHintLabel = New-Object System.Windows.Forms.Label
$ApplyHintLabel.Location = New-Object System.Drawing.Point(32, 432)
$ApplyHintLabel.Size = New-Object System.Drawing.Size(238, 38)
$ApplyHintLabel.Text = 'This applies the cached image using the native Windows wallpaper API.'
$ApplyHintLabel.ForeColor = $script:TextMutedOnDarkColor
$ApplyHintLabel.BackColor = [System.Drawing.Color]::Transparent
$ApplyHintLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
$LeftPanel.Controls.Add($ApplyHintLabel)

$StatusCard = New-Object System.Windows.Forms.Panel
$StatusCard.Size = New-Object System.Drawing.Size(244, 164)
$StatusCard.Location = New-Object System.Drawing.Point(28, 546)
$StatusCard.BackColor = $script:SidebarCardColor
Enable-RoundedCorners -Control $StatusCard -Radius 22
$LeftPanel.Controls.Add($StatusCard)

$StatusAccentBar = New-Object System.Windows.Forms.Panel
$StatusAccentBar.Size = New-Object System.Drawing.Size(52, 4)
$StatusAccentBar.Location = New-Object System.Drawing.Point(18, 16)
$StatusAccentBar.BackColor = $script:AccentBlueColor
Enable-RoundedCorners -Control $StatusAccentBar -Radius 2
$StatusCard.Controls.Add($StatusAccentBar)

$StatusTitleLabel = New-Object System.Windows.Forms.Label
$StatusTitleLabel.AutoSize = $true
$StatusTitleLabel.Location = New-Object System.Drawing.Point(18, 30)
$StatusTitleLabel.Text = 'LIVE STATUS'
$StatusTitleLabel.ForeColor = $script:TextMutedOnDarkColor
$StatusTitleLabel.BackColor = [System.Drawing.Color]::Transparent
$StatusTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$StatusCard.Controls.Add($StatusTitleLabel)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Point(18, 58)
$StatusLabel.Size = New-Object System.Drawing.Size(208, 72)
$StatusLabel.Text = 'Ready to pull a wallpaper.'
$StatusLabel.ForeColor = $script:TextOnDarkColor
$StatusLabel.BackColor = [System.Drawing.Color]::Transparent
$StatusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10.2)
$StatusCard.Controls.Add($StatusLabel)

$HeaderPanel = New-Object System.Windows.Forms.Panel
$HeaderPanel.Dock = [System.Windows.Forms.DockStyle]::Top
$HeaderPanel.Height = 78
$HeaderPanel.BackColor = $script:SurfaceColor
$PreviewHost.Controls.Add($HeaderPanel)

$HeaderBadgePanel = New-Object System.Windows.Forms.Panel
$HeaderBadgePanel.Dock = [System.Windows.Forms.DockStyle]::Right
$HeaderBadgePanel.Width = 260
$HeaderBadgePanel.BackColor = $script:SurfaceColor
$HeaderPanel.Controls.Add($HeaderBadgePanel)

$HeaderTextPanel = New-Object System.Windows.Forms.Panel
$HeaderTextPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$HeaderTextPanel.BackColor = $script:SurfaceColor
$HeaderPanel.Controls.Add($HeaderTextPanel)

$HeaderTitleLabel = New-Object System.Windows.Forms.Label
$HeaderTitleLabel.Dock = [System.Windows.Forms.DockStyle]::Top
$HeaderTitleLabel.Height = 40
$HeaderTitleLabel.Text = 'Bring more color to the desktop'
$HeaderTitleLabel.ForeColor = $script:TextPrimaryColor
$HeaderTitleLabel.BackColor = [System.Drawing.Color]::Transparent
$HeaderTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 25)
$HeaderTextPanel.Controls.Add($HeaderTitleLabel)

$HeaderSubtitleLabel = New-Object System.Windows.Forms.Label
$HeaderSubtitleLabel.Dock = [System.Windows.Forms.DockStyle]::Top
$HeaderSubtitleLabel.Height = 24
$HeaderSubtitleLabel.Text = 'Preview a fresh pull from Wallhaven, then apply it when it feels right.'
$HeaderSubtitleLabel.ForeColor = $script:TextSecondaryColor
$HeaderSubtitleLabel.BackColor = [System.Drawing.Color]::Transparent
$HeaderSubtitleLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10.5)
$HeaderTextPanel.Controls.Add($HeaderSubtitleLabel)

$WallhavenBadge = New-Object System.Windows.Forms.Label
$WallhavenBadge.AutoSize = $true
$WallhavenBadge.Location = New-Object System.Drawing.Point(26, 16)
$WallhavenBadge.Text = ' Wallhaven '
$WallhavenBadge.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#E4F5FF')
$WallhavenBadge.ForeColor = $script:TextPrimaryColor
$WallhavenBadge.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$HeaderBadgePanel.Controls.Add($WallhavenBadge)

$SearchBadge = New-Object System.Windows.Forms.Label
$SearchBadge.AutoSize = $true
$SearchBadge.Location = New-Object System.Drawing.Point(130, 16)
$SearchBadge.Text = ' Canada nature '
$SearchBadge.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#E8F6DD')
$SearchBadge.ForeColor = $script:TextPrimaryColor
$SearchBadge.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$HeaderBadgePanel.Controls.Add($SearchBadge)

$PreviewShell = New-Object System.Windows.Forms.Panel
$PreviewShell.Dock = [System.Windows.Forms.DockStyle]::Fill
$PreviewShell.BackColor = $script:SoftBorderColor
$PreviewHost.Controls.Add($PreviewShell)

$PreviewBorder = New-Object System.Windows.Forms.Panel
$PreviewBorder.Dock = [System.Windows.Forms.DockStyle]::Fill
$PreviewBorder.Padding = New-Object System.Windows.Forms.Padding(1)
$PreviewBorder.BackColor = [System.Drawing.Color]::White
$PreviewShell.Controls.Add($PreviewBorder)

$PreviewHeaderStrip = New-Object System.Windows.Forms.Panel
$PreviewHeaderStrip.Dock = [System.Windows.Forms.DockStyle]::Top
$PreviewHeaderStrip.Height = 48
$PreviewHeaderStrip.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#EEF5F3')
$PreviewBorder.Controls.Add($PreviewHeaderStrip)

$PreviewTitleLabel = New-Object System.Windows.Forms.Label
$PreviewTitleLabel.AutoSize = $true
$PreviewTitleLabel.Location = New-Object System.Drawing.Point(18, 15)
$PreviewTitleLabel.Text = 'Preview'
$PreviewTitleLabel.ForeColor = $script:TextPrimaryColor
$PreviewTitleLabel.BackColor = [System.Drawing.Color]::Transparent
$PreviewTitleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
$PreviewHeaderStrip.Controls.Add($PreviewTitleLabel)

$PreviewMetaInfoLabel = New-Object System.Windows.Forms.Label
$PreviewMetaInfoLabel.AutoSize = $true
$PreviewMetaInfoLabel.ForeColor = $script:TextSecondaryColor
$PreviewMetaInfoLabel.BackColor = [System.Drawing.Color]::Transparent
$PreviewMetaInfoLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9.5)
$PreviewMetaInfoLabel.Text = 'Standing by'
$PreviewHeaderStrip.Controls.Add($PreviewMetaInfoLabel)
$PreviewHeaderStrip.Add_Resize({
    $PreviewMetaInfoLabel.Left = $PreviewHeaderStrip.ClientSize.Width - $PreviewMetaInfoLabel.Width - 18
    $PreviewMetaInfoLabel.Top = 15
})

$PreviewInfoBar = New-Object System.Windows.Forms.Panel
$PreviewInfoBar.Dock = [System.Windows.Forms.DockStyle]::Bottom
$PreviewInfoBar.Height = 58
$PreviewInfoBar.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#F3F8F6')
$PreviewBorder.Controls.Add($PreviewInfoBar)

$MetaResolutionLabel = New-Object System.Windows.Forms.Label
$MetaResolutionLabel.AutoSize = $true
$MetaResolutionLabel.Location = New-Object System.Drawing.Point(18, 20)
$MetaResolutionLabel.ForeColor = $script:TextSecondaryColor
$MetaResolutionLabel.BackColor = [System.Drawing.Color]::Transparent
$MetaResolutionLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9.5)
$PreviewInfoBar.Controls.Add($MetaResolutionLabel)

$MetaIdLabel = New-Object System.Windows.Forms.Label
$MetaIdLabel.AutoSize = $true
$MetaIdLabel.Location = New-Object System.Drawing.Point(250, 20)
$MetaIdLabel.ForeColor = $script:TextSecondaryColor
$MetaIdLabel.BackColor = [System.Drawing.Color]::Transparent
$MetaIdLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9.5)
$PreviewInfoBar.Controls.Add($MetaIdLabel)

$MetaTypeLabel = New-Object System.Windows.Forms.Label
$MetaTypeLabel.AutoSize = $true
$MetaTypeLabel.Location = New-Object System.Drawing.Point(480, 20)
$MetaTypeLabel.ForeColor = $script:TextSecondaryColor
$MetaTypeLabel.BackColor = [System.Drawing.Color]::Transparent
$MetaTypeLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9.5)
$PreviewInfoBar.Controls.Add($MetaTypeLabel)

$MetaStyleLabel = New-Object System.Windows.Forms.Label
$MetaStyleLabel.AutoSize = $true
$MetaStyleLabel.Location = New-Object System.Drawing.Point(670, 20)
$MetaStyleLabel.ForeColor = $script:TextSecondaryColor
$MetaStyleLabel.BackColor = [System.Drawing.Color]::Transparent
$MetaStyleLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9.5)
$PreviewInfoBar.Controls.Add($MetaStyleLabel)

$PreviewCanvas = New-Object System.Windows.Forms.Panel
$PreviewCanvas.Dock = [System.Windows.Forms.DockStyle]::Fill
$PreviewCanvas.Padding = New-Object System.Windows.Forms.Padding(18)
$PreviewCanvas.BackColor = [System.Drawing.Color]::White
$PreviewBorder.Controls.Add($PreviewCanvas)

$PictureBox = New-Object System.Windows.Forms.PictureBox
$PictureBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$PictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$PictureBox.BackColor = [System.Drawing.Color]::White
$PreviewCanvas.Controls.Add($PictureBox)

$PlaceholderLabel = New-Object System.Windows.Forms.Label
$PlaceholderLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$PlaceholderLabel.Text = "Random picture pulled from the feed`r`n`r`nUse New Picture to rotate through the Wallhaven results."
$PlaceholderLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$PlaceholderLabel.ForeColor = $script:TextSecondaryColor
$PlaceholderLabel.BackColor = [System.Drawing.Color]::White
$PlaceholderLabel.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 18)
$PreviewCanvas.Controls.Add($PlaceholderLabel)
$PlaceholderLabel.BringToFront()
$config = Get-AppConfig
$script:State.SearchUrl = [string]$config.WallhavenSearchUrl
$script:State.ApiUrl = Convert-WallhavenSearchUrlToApiUrl -SearchUrl $script:State.SearchUrl
$script:State.WallpaperStyle = [string]$config.WallpaperStyle
Reset-WallpaperDetails
Update-ActionButtons

$NewPictureButton.Add_Click({ Load-NewWallpaper })
$SetWallpaperButton.Add_Click({
    if (-not $script:State.CurrentFilePath) {
        Set-Status -Message 'Pull an image first before setting the wallpaper.' -Color $script:ErrorColor
        return
    }
    try {
        Set-BusyState -IsBusy $true
        Set-Status -Message 'Applying the wallpaper in Windows...' -Color $script:NeutralColor
        Set-DesktopWallpaper -Path $script:State.CurrentFilePath
        Set-Status -Message 'Wallpaper applied.' -Color $script:SuccessColor
    }
    catch {
        Set-Status -Message $_.Exception.Message -Color $script:ErrorColor
    }
    finally {
        Set-BusyState -IsBusy $false
    }
})

$MainForm.Add_Shown({ Load-NewWallpaper })

if ($SmokeTest) {
    Initialize-CacheDirectory
    Write-Output 'Smoke test passed.'
    return
}

[void]$MainForm.ShowDialog()


