# PixelLock from the command line

[Svenska](kommandoraden.md)

PixelLock can take screenshots without its window: from PowerShell, a script, a scheduled task or
an AI tool such as Claude Code. This page shows what it can do, with recipes you can copy.

- [The basics](#the-basics)
- [Recipe: a screenshot of every tab in a report](#recipe-a-screenshot-of-every-tab-in-a-report)
- [More recipes](#more-recipes)
- [AI tools: Claude Code and Claude Desktop](#ai-tools-claude-code-and-claude-desktop)
- [All options](#all-options)
- [When something goes wrong](#when-something-goes-wrong)

## The basics

Open PowerShell (Start menu, type *PowerShell*) and type:

```powershell
pixellock capture | Out-String
```

PixelLock takes a screenshot of all your screens, saves it where Settings > Saving says, and
prints where the file went. Nothing opens on the screen.

**Always end with `| Out-String`.** PixelLock is a Windows app, not a console program, so without
it PowerShell neither waits for the screenshot nor shows the file's path. With it, PowerShell
waits, and in a script `$LASTEXITCODE` tells you whether it worked (0 means saved).

A few more:

```powershell
# Only the main screen, into a file you name
pixellock capture --primary --out C:\Shots\start.png | Out-String

# One window on its own, even with other windows over it
pixellock capture --window Excel | Out-String

# Wait 3 seconds first, for a menu or a tooltip to show
pixellock capture --window Excel --delay 3 | Out-String

# The window you are working in: you have 3 seconds to click it
pixellock capture --active --delay 3 | Out-String

# Which windows are open, and what are they called?
pixellock windows | Out-String
```

`--window` takes any part of the window's title or app name, in any case: `Excel`,
`"Sales report"`, `Edge`. If several windows match, the one in front is taken. A minimized window
cannot be captured.

`--active` needs no title: it takes the window that is active when the screenshot is taken. Typed
in PowerShell that is PowerShell itself, so give yourself a `--delay` to click the app, or let a
script switch to it first (as the recipe below does with `AppActivate`).

`pixellock` is found by name once PixelLock is installed, from the Microsoft Store or the
download on this site. If PowerShell says it is not found, open a new PowerShell window, or use
the full path, `$env:LOCALAPPDATA\PixelLock\pixellock.exe`.

## Recipe: a screenshot of every tab in a report

You have a report with seven tabs (Excel sheets, browser tabs, pages in an app) and want one
screenshot of each, named so they stay in order.

1. Download [capture-tabs.ps1](recipes/capture-tabs.ps1) (on that page: **Download raw file**)
   and save it, for example in `Documents`.
2. Open the report on its **first** tab.
3. Open PowerShell in the folder with the script and run:

```powershell
.\capture-tabs.ps1 -Window "Sales report" -Tabs 7
```

The script brings the report to the front, captures the window, presses Ctrl+Page Down to go to
the next tab and repeats until all seven are done. You get `report-tab-1.png` to
`report-tab-7.png` in `Pictures\PixelLock\Report`, and the folder opens.

Ctrl+Page Down moves to the next sheet in Excel and the next tab in Edge, Chrome and Firefox.

| You want                                       | Add                                                 |
| ---------------------------------------------- | --------------------------------------------------- |
| Another folder                                 | `-Folder C:\Reports\November`                       |
| Other file names (`budget-1.png` ...)          | `-Name budget`                                      |
| All tabs in one long image as well             | `-Combine` (gives `report-tab-all.png`)             |
| More time for a slow report to draw            | `-Wait 5` (seconds; 2 if you leave it out)          |
| Another key for the next tab                   | `-NextKey "^{TAB}"` (Ctrl+Tab), `-NextKey "{RIGHT}"` |
| To click the tabs yourself                     | `-Manual`                                           |

Use **`-Manual`** when the tabs are links on a web page, as in Power BI in the browser, where
Ctrl+Page Down would move to another browser tab instead. The script then asks you to open each
tab and press Enter in PowerShell. The report may stay behind PowerShell: PixelLock captures its
window on its own.

```powershell
.\capture-tabs.ps1 -Window "Power BI" -Tabs 7 -Manual -Combine
```

If Windows says running scripts is disabled, run this once, then try again:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
Unblock-File .\capture-tabs.ps1
```

**The same thing by hand, in a few lines**, if you would rather see every step:

```powershell
$folder = "$HOME\Pictures\PixelLock\Report"
$shell = New-Object -ComObject WScript.Shell
for ($tab = 1; $tab -le 7; $tab++) {
    pixellock capture --window "Sales report" --delay 2 --out "$folder\report-tab-$tab.png" | Out-String
    [void] $shell.AppActivate("Sales report")
    $shell.SendKeys("^{PGDN}")
}
```

## More recipes

**A screenshot every 10 minutes** (for example of a dashboard), until you press Ctrl+C:

```powershell
while ($true) {
    pixellock capture --window "Dashboard" --out "$HOME\Pictures\Dashboard" | Out-String
    Start-Sleep -Seconds 600
}
```

With a folder as `--out`, each file is named with your file name pattern from Settings > Saving,
so they do not overwrite each other.

**Every morning at 8**, with Windows Task Scheduler: create a basic task, choose *Start a
program*, program `pixellock`, arguments `capture --primary --out C:\Shots`.

**The same area each time**, in physical pixels from the top-left corner of the main screen:

```powershell
pixellock capture --region 0,0,1280,720 --out C:\Shots\area.png | Out-String
```

**Straight to the clipboard** as well, to paste into a chat or a mail:

```powershell
pixellock capture --window Excel --clipboard | Out-String
```

**Use the path in the rest of a script**:

```powershell
$path = (pixellock capture --primary | Out-String).Trim()
if ($LASTEXITCODE -ne 0) { throw "No screenshot" }
Copy-Item $path "\\server\share\reports\"
```

**In a series by hand**: take one screenshot each time you press Enter.

```powershell
for ($n = 1; $true; $n++) {
    Read-Host "Press Enter for screenshot $n (Ctrl+C to stop)" | Out-Null
    pixellock capture --primary --delay 1 --out "$HOME\Pictures\Steps\step-$n.png" | Out-String
}
```

## AI tools: Claude Code and Claude Desktop

**Claude Code** can run `pixellock capture` itself, like any other command. Ask it, for example,
"take a screenshot of the Excel window with pixellock and tell me what the chart shows". It
reads the printed path and opens the file.

**Claude Desktop and Claude Code** can also read your PixelLock history through its MCP server:
list, search and look at the screenshots you have taken, and read their text. First turn on
**Settings > AI > Let Claude read your captures** in PixelLock, then add the server.

Claude Code:

```powershell
claude mcp add pixellock -- pixellock mcp
```

Claude Desktop: Settings > Developer > Edit Config, add this to `claude_desktop_config.json` and
restart Claude Desktop:

```json
{
  "mcpServers": {
    "pixellock": { "command": "pixellock", "args": ["mcp"] }
  }
}
```

The MCP server can only read. It cannot take new screenshots; for that, Claude Code uses
`pixellock capture` as above.

## All options

```text
pixellock capture [WHAT] [--delay SECONDS] [--out PATH] [--clipboard]
pixellock windows
pixellock mcp
pixellock --help
```

| Option              | What it does                                                                                       |
| ------------------- | -------------------------------------------------------------------------------------------------- |
| `--all`             | Every screen as one image (what you get with no WHAT)                                              |
| `--primary`         | The main screen                                                                                    |
| `--monitor N`       | Screen N, counted from 1                                                                           |
| `--region X,Y,W,H`  | A rectangle in physical pixels; X and Y may be negative for a screen left of or above the main one |
| `--window TEXT`     | The front-most window with TEXT in its title or app name, on its own                               |
| `--active`          | The window you are working in, on its own                                                          |
| `--delay SECONDS`   | Wait first: 0 to 600 seconds, `2`, `0.5` and `0,5` all work                                        |
| `--out PATH`        | A file ending in `.png` or `.jpg`, or a folder. Without it: where Settings > Saving says            |
| `--clipboard`       | Also copy the screenshot                                                                           |

`pixellock windows` lists the open windows, front-most first: the app name, a tab and the title.

Exit codes in `$LASTEXITCODE`: `0` saved, `1` the screenshot or the save failed (or no window
matched `--window`, or no window was active for `--active`), `2` something is wrong in the command.

## When something goes wrong

| What you see                                          | What to do                                                                                                                         |
| ----------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `pixellock` is not recognized                         | Open a new PowerShell window. Still not? Use `$env:LOCALAPPDATA\PixelLock\pixellock.exe`.                                           |
| Nothing is printed and the script races ahead         | Add `\| Out-String` to the end of the `pixellock` line.                                                                            |
| `No open window has "..." in its title or app name`   | Run `pixellock windows \| Out-String` and use a word from the title it shows. Restore the window if it is minimized.             |
| The wrong window was captured                         | Use more of the title, for example `"Sales report.xlsx"` instead of `Excel`.                                                       |
| `--active` took PowerShell                            | Add `--delay 3` and click the app, or switch to it in the script before the `pixellock` line.                                    |
| `No window is active`                                 | The desktop was active. Click the window, or switch to it in the script, before the screenshot.                                   |
| The screenshot shows the previous tab                 | Give the report more time: `-Wait 4` in the script, or a larger `--delay`.                                                          |
| The key went to the wrong place                       | Use `-Manual`, or another `-NextKey` that your app uses for the next tab.                                                           |
| Running scripts is disabled on this system            | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, then `Unblock-File .\capture-tabs.ps1`.                                     |

Missing something? Tell us with **Send feedback** in PixelLock, or
[open an issue](https://github.com/Beastlander/pixellock-public/issues/new/choose).
