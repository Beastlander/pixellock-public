# PixelLock från kommandoraden

[English](command-line.md)

PixelLock kan ta skärmbilder utan sitt fönster: från PowerShell, ett skript, en schemalagd
aktivitet eller ett AI-verktyg som Claude Code. Här står vad det kan, med recept att kopiera.

- [Grunderna](#grunderna)
- [Recept: en skärmbild av varje flik i en rapport](#recept-en-skärmbild-av-varje-flik-i-en-rapport)
- [Fler recept](#fler-recept)
- [AI-verktyg: Claude Code och Claude Desktop](#ai-verktyg-claude-code-och-claude-desktop)
- [Alla val](#alla-val)
- [Om något går fel](#om-något-går-fel)

## Grunderna

Öppna PowerShell (Start-menyn, skriv *PowerShell*) och skriv:

```powershell
pixellock capture
```

PixelLock tar en skärmbild av alla dina skärmar, sparar den där Inställningar > Spara säger och
skriver ut var filen hamnade. Inget öppnas på skärmen.

PowerShell väntar tills skärmbilden är sparad, och i ett skript säger `$LASTEXITCODE` om det gick
bra (0 betyder sparad). Det fungerar likadant i Kommandotolken, med `%ERRORLEVEL%`. (PixelLock
0.5.0 och äldre behövde `| Out-String` sist på raden för det.)

Några till:

```powershell
# Bara huvudskärmen, till en fil du namnger
pixellock capture --primary --out C:\Bilder\start.png

# Ett fönster för sig, även med andra fönster över det
pixellock capture --window Excel

# Vänta 3 sekunder först, så att en meny eller en knappbeskrivning hinner visas
pixellock capture --window Excel --delay 3

# Fönstret du arbetar i: du har 3 sekunder på dig att klicka på det
pixellock capture --active --delay 3

# Fönstret du arbetar i: du har 3 sekunder på dig att klicka på det
pixellock capture --active --delay 3 | Out-String

# Vilka fönster är öppna, och vad heter de?
pixellock windows
```

`--window` tar valfri del av fönstrets titel eller programmets namn, stora eller små bokstäver:
`Excel`, `"Försäljningsrapport"`, `Edge`. Passar flera fönster tas det som ligger främst. Ett
minimerat fönster går inte att ta.

`--active` behöver ingen titel: det tar fönstret som är aktivt när skärmbilden tas. Skrivet i
PowerShell är det PowerShell självt, så ge dig ett `--delay` att klicka på programmet, eller låt
ett skript byta till det först (som receptet nedan gör med `AppActivate`).

`pixellock` hittas med bara namnet när PixelLock är installerat, från Microsoft Store eller
nedladdningen på den här webbplatsen. Säger PowerShell att det inte finns: öppna ett nytt
PowerShell-fönster, eller använd hela sökvägen, `$env:LOCALAPPDATA\PixelLock\pixellock-cli.exe`.

## Recept: en skärmbild av varje flik i en rapport

Du har en rapport med sju flikar (blad i Excel, flikar i webbläsaren, sidor i ett program) och
vill ha en skärmbild av varje, med namn som håller ordningen.

1. Ladda ner [capture-tabs.ps1](recipes/capture-tabs.ps1) (på den sidan: **Download raw
   file**) och spara den, till exempel i `Dokument`.
2. Öppna rapporten på **första** fliken.
3. Öppna PowerShell i mappen med skriptet och kör:

```powershell
.\capture-tabs.ps1 -Window "Försäljningsrapport" -Tabs 7
```

Skriptet tar fram rapporten, tar en bild av fönstret, trycker Ctrl+Page Down för att gå till
nästa flik och fortsätter tills alla sju är klara. Du får `report-tab-1.png` till
`report-tab-7.png` i `Bilder\PixelLock\Report`, och mappen öppnas.

Ctrl+Page Down går till nästa blad i Excel och nästa flik i Edge, Chrome och Firefox.

| Du vill                                         | Lägg till                                              |
| ----------------------------------------------- | ------------------------------------------------------ |
| En annan mapp                                   | `-Folder C:\Rapporter\November`                        |
| Andra filnamn (`budget-1.png` ...)              | `-Name budget`                                         |
| Alla flikar i en lång bild också                | `-Combine` (ger `report-tab-all.png`)                  |
| Mer tid för en långsam rapport att ritas        | `-Wait 5` (sekunder; 2 om du inte anger något)         |
| En annan tangent för nästa flik                 | `-NextKey "^{TAB}"` (Ctrl+Tab), `-NextKey "{RIGHT}"`    |
| Klicka på flikarna själv                        | `-Manual`                                              |

Använd **`-Manual`** när flikarna är länkar på en webbsida, som i Power BI i webbläsaren, där
Ctrl+Page Down i stället byter webbläsarflik. Skriptet ber dig då öppna varje flik och trycka
Enter i PowerShell. Rapporten får ligga bakom PowerShell: PixelLock tar dess fönster för sig.

```powershell
.\capture-tabs.ps1 -Window "Power BI" -Tabs 7 -Manual -Combine
```

Säger Windows att skript är avstängda: kör detta en gång och försök igen:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
Unblock-File .\capture-tabs.ps1
```

**Samma sak för hand, på några rader**, om du hellre ser varje steg:

```powershell
$folder = "$HOME\Pictures\PixelLock\Report"
$shell = New-Object -ComObject WScript.Shell
for ($tab = 1; $tab -le 7; $tab++) {
    pixellock capture --window "Försäljningsrapport" --delay 2 --out "$folder\report-tab-$tab.png"
    [void] $shell.AppActivate("Försäljningsrapport")
    $shell.SendKeys("^{PGDN}")
}
```

## Fler recept

**En skärmbild var tionde minut** (till exempel av en översiktspanel), tills du trycker Ctrl+C:

```powershell
while ($true) {
    pixellock capture --window "Dashboard" --out "$HOME\Pictures\Dashboard"
    Start-Sleep -Seconds 600
}
```

Med en mapp som `--out` får varje fil ett namn efter ditt filnamnsmönster i Inställningar >
Spara, så att de inte skriver över varandra.

**Varje morgon klockan 8**, med Schemaläggaren i Windows: skapa en enkel aktivitet, välj
*Starta ett program*, program `pixellock`, argument `capture --primary --out C:\Bilder`.

**Samma område varje gång**, i fysiska pixlar från huvudskärmens övre vänstra hörn:

```powershell
pixellock capture --region 0,0,1280,720 --out C:\Bilder\omrade.png
```

**Direkt till urklipp** också, för att klistra in i en chatt eller ett mejl:

```powershell
pixellock capture --window Excel --clipboard
```

**Använd sökvägen i resten av ett skript**:

```powershell
$path = pixellock capture --primary
if ($LASTEXITCODE -ne 0) { throw "Ingen skärmbild" }
Copy-Item $path "\\server\delad\rapporter\"
```

**En serie för hand**: en skärmbild varje gång du trycker Enter.

```powershell
for ($n = 1; $true; $n++) {
    Read-Host "Tryck Enter för skärmbild $n (Ctrl+C avslutar)" | Out-Null
    pixellock capture --primary --delay 1 --out "$HOME\Pictures\Steg\steg-$n.png"
}
```

## AI-verktyg: Claude Code och Claude Desktop

**Claude Code** kan köra `pixellock capture` själv, som vilket kommando som helst. Be till
exempel om "ta en skärmbild av Excel-fönstret med pixellock och berätta vad diagrammet visar".
Det läser den utskrivna sökvägen och öppnar filen.

**Claude Desktop och Claude Code** kan också läsa din PixelLock-historik genom MCP-servern:
lista, söka och titta på skärmbilderna du har tagit, och läsa texten i dem. Slå först på
**Inställningar > AI > Låt Claude läsa dina skärmbilder** i PixelLock och lägg sedan till
servern.

Claude Code:

```powershell
claude mcp add pixellock -- pixellock mcp
```

Claude Desktop: Settings > Developer > Edit Config, lägg till detta i
`claude_desktop_config.json` och starta om Claude Desktop:

```json
{
  "mcpServers": {
    "pixellock": { "command": "pixellock", "args": ["mcp"] }
  }
}
```

MCP-servern kan bara läsa. Den kan inte ta nya skärmbilder; till det använder Claude Code
`pixellock capture` som ovan.

## Alla val

```text
pixellock capture [VAD] [--delay SEKUNDER] [--out SÖKVÄG] [--clipboard]
pixellock windows
pixellock mcp
pixellock --help
```

| Val                 | Vad det gör                                                                                                  |
| ------------------- | ------------------------------------------------------------------------------------------------------------ |
| `--all`             | Alla skärmar som en bild (det du får utan VAD)                                                               |
| `--primary`         | Huvudskärmen                                                                                                 |
| `--monitor N`       | Skärm N, räknat från 1                                                                                       |
| `--region X,Y,W,H`  | En rektangel i fysiska pixlar; X och Y kan vara negativa för en skärm till vänster om eller ovanför huvudskärmen |
| `--window TEXT`     | Det främsta fönstret med TEXT i titeln eller programnamnet, för sig                                          |
| `--active`          | Fönstret du arbetar i, för sig                                                                               |
| `--delay SEKUNDER`  | Vänta först: 0 till 600 sekunder, `2`, `0.5` och `0,5` fungerar                                               |
| `--out SÖKVÄG`      | En fil som slutar på `.png` eller `.jpg`, eller en mapp. Utan: där Inställningar > Spara säger               |
| `--clipboard`       | Kopiera skärmbilden också                                                                                    |

`pixellock windows` listar de öppna fönstren, det främsta först: programnamnet, en tabb och
titeln.

Slutkoder i `$LASTEXITCODE`: `0` sparad, `1` skärmbilden eller sparandet misslyckades (eller
inget fönster passade `--window`, eller inget fönster var aktivt för `--active`), `2` något är fel i kommandot.

## Om något går fel

| Det du ser                                             | Gör så här                                                                                                                  |
| ------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------- |
| `pixellock` känns inte igen                            | Öppna ett nytt PowerShell-fönster. Fortfarande inte? Använd `$env:LOCALAPPDATA\PixelLock\pixellock-cli.exe`.                |
| Inget skrivs ut och skriptet rusar i väg               | Uppdatera PixelLock. En äldre version behöver `\| Out-String` sist på raden med `pixellock`.                               |
| `No open window has "..." in its title or app name`    | Kör `pixellock windows` och använd ett ord ur titeln som visas. Återställ fönstret om det är minimerat.       |
| Fel fönster kom med                                    | Använd mer av titeln, till exempel `"Försäljningsrapport.xlsx"` i stället för `Excel`.                                       |
| `--active` tog PowerShell                              | Lägg till `--delay 3` och klicka på programmet, eller byt till det i skriptet före raden med `pixellock`.                   |
| `No window is active`                                  | Skrivbordet var aktivt. Klicka på fönstret, eller byt till det i skriptet, före skärmbilden.                                 |
| Skärmbilden visar förra fliken                         | Ge rapporten mer tid: `-Wait 4` i skriptet, eller ett större `--delay`.                                                     |
| Tangenten hamnade fel                                  | Använd `-Manual`, eller en annan `-NextKey` som ditt program använder för nästa flik.                                        |
| Skript är avstängda i det här systemet                 | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, sedan `Unblock-File .\capture-tabs.ps1`.                              |

Saknar du något? Säg till med **Skicka feedback** i PixelLock, eller
[skapa ett ärende](https://github.com/Beastlander/pixellock-public/issues/new/choose).
