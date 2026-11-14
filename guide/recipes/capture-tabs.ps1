<#
.SYNOPSIS
    Takes a screenshot of every tab in a report: one file per tab, and optionally one image of
    them all.

.DESCRIPTION
    Open the report on its first tab, then run this script. It brings the report window to the
    front, waits for it to draw, captures the window alone with PixelLock, moves to the next tab
    with a key press and repeats. The files are named report-tab-1.png, report-tab-2.png and so
    on, in the folder you choose.

    The next-tab key is Ctrl+Page Down by default. It moves to the next sheet in Excel and to the
    next tab in Edge, Chrome and Firefox. For an app with another key, give it with -NextKey. For
    a report whose tabs are links on a web page (such as Power BI in the browser), use -Manual:
    you click each tab yourself and press Enter here.

    Needs PixelLock 0.4 or later (for --window and --delay). Works in Windows PowerShell 5.1 and
    PowerShell 7 on Windows.

.PARAMETER Window
    Text in the report window's title, for example "Sales report" or "Excel". Not case
    sensitive. Run `pixellock windows | Out-String` to see the titles of the open windows.

.PARAMETER Tabs
    How many tabs to capture. 7 if you leave it out.

.PARAMETER Folder
    Where to save the screenshots. Pictures\PixelLock\Report if you leave it out.

.PARAMETER Name
    The start of each file name. "report-tab" gives report-tab-1.png, report-tab-2.png ...

.PARAMETER NextKey
    The key that moves to the next tab, written for SendKeys: ^ is Ctrl, + is Shift, % is Alt,
    {PGDN} is Page Down, {TAB} is Tab, {RIGHT} is the right arrow. "^{PGDN}" if you leave it out.

.PARAMETER Wait
    Seconds to let each tab draw before it is captured. 2 if you leave it out; give a slow
    report more.

.PARAMETER Manual
    Do not press any key: you open each tab yourself and press Enter in this window. The report
    may stay behind this window, since PixelLock captures it alone.

.PARAMETER Combine
    Also put all the tabs under each other in one image, <Name>-all.png.

.EXAMPLE
    .\capture-tabs.ps1 -Window "Sales report" -Tabs 7

.EXAMPLE
    .\capture-tabs.ps1 -Window "Budget.xlsx" -Tabs 7 -Folder C:\Reports\November -Combine

.EXAMPLE
    .\capture-tabs.ps1 -Window "Power BI" -Tabs 7 -Manual
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Window,

    [ValidateRange(1, 100)]
    [int] $Tabs = 7,

    [string] $Folder = (Join-Path ([Environment]::GetFolderPath('MyPictures')) 'PixelLock\Report'),

    [string] $Name = 'report-tab',

    [string] $NextKey = '^{PGDN}',

    [ValidateRange(0, 600)]
    [double] $Wait = 2,

    [switch] $Manual,

    [switch] $Combine
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command pixellock -ErrorAction SilentlyContinue)) {
    throw 'pixellock was not found. Install PixelLock, then open a new PowerShell window.'
}

# Brings the report window to the front, so the next-tab key reaches it.
function Show-Report {
    $process = Get-Process |
        Where-Object { $_.MainWindowTitle.IndexOf($Window, [StringComparison]::OrdinalIgnoreCase) -ge 0 } |
        Select-Object -First 1
    if (-not $process) {
        throw "No window has '$Window' in its title. Run 'pixellock windows | Out-String' to see the open windows."
    }
    $shell = New-Object -ComObject WScript.Shell
    [void] $shell.AppActivate($process.Id)
    Start-Sleep -Milliseconds 300
}

# Captures the report window alone into $File after waiting $Wait seconds. Piping to Out-String
# makes PowerShell wait for PixelLock and set $LASTEXITCODE.
function Save-Tab([string] $File) {
    $delay = $Wait.ToString([Globalization.CultureInfo]::InvariantCulture)
    $saved = pixellock capture --window $Window --delay $delay --out $File | Out-String
    if ($LASTEXITCODE -ne 0) {
        throw "PixelLock could not capture '$Window' (exit code $LASTEXITCODE)."
    }
    $saved.Trim()
}

# Puts the images under each other, as wide as the widest, on white, and saves it as a PNG.
function Join-Images([string[]] $Files, [string] $Target) {
    Add-Type -AssemblyName System.Drawing
    $images = @($Files | ForEach-Object { [System.Drawing.Image]::FromFile($_) })
    try {
        $width = ($images | Measure-Object -Property Width -Maximum).Maximum
        $height = ($images | Measure-Object -Property Height -Sum).Sum
        $sheet = New-Object System.Drawing.Bitmap ([int] $width), ([int] $height)
        $graphics = [System.Drawing.Graphics]::FromImage($sheet)
        try {
            $graphics.Clear([System.Drawing.Color]::White)
            $top = 0
            foreach ($image in $images) {
                $graphics.DrawImage($image, 0, $top, $image.Width, $image.Height)
                $top += $image.Height
            }
            $sheet.Save($Target, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $graphics.Dispose()
            $sheet.Dispose()
        }
    }
    finally {
        $images | ForEach-Object { $_.Dispose() }
    }
}

New-Item -ItemType Directory -Force -Path $Folder | Out-Null
$files = @()
for ($tab = 1; $tab -le $Tabs; $tab++) {
    if ($Manual) {
        Read-Host "Open tab $tab of $Tabs in the report, then press Enter here" | Out-Null
    }
    else {
        Show-Report
        if ($tab -gt 1) {
            (New-Object -ComObject WScript.Shell).SendKeys($NextKey)
        }
    }
    $file = Join-Path $Folder "$Name-$tab.png"
    $files += Save-Tab $file
    Write-Host "Tab $tab of ${Tabs}: $file"
}

if ($Combine) {
    $all = Join-Path $Folder "$Name-all.png"
    Join-Images $files $all
    Write-Host "All tabs in one image: $all"
}

Invoke-Item $Folder
