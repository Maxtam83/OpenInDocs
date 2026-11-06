<#
  Install.ps1 - installs OpenInDocs for the current user (no administrator rights needed).
  Right-click > "Run with PowerShell".   Uninstall:  powershell -ExecutionPolicy Bypass -File Install.ps1 -Uninstall

  Installe OpenInDocs pour l'utilisateur courant (sans droits administrateur).
  Clic droit > "Exécuter avec PowerShell".   Désinstaller : Install.ps1 -Uninstall

  What it does / Ce que ça fait:
    1. compiles the 3 launchers with the C# compiler shipped with Windows;
    2. extracts the Docs / Sheets / Slides icons from YOUR Google Drive for desktop installation
       (nothing from Google is stored in this repository);
    3. registers the launchers with Windows (file types, icons, "Open with" list);
    4. removes older versions (OuvrirDansDocs v1/v2).
#>
param([switch]$Uninstall)

$ErrorActionPreference = 'Stop'
$Root    = Split-Path -Parent $PSScriptRoot
$Dest    = Join-Path $env:LOCALAPPDATA 'OpenInDocs'
$OldDest = Join-Path $env:LOCALAPPDATA 'OuvrirDansDocs'
$Classes = 'HKCU:\Software\Classes'
$RegApps = 'HKCU:\Software\RegisteredApplications'
$Fr      = ([System.Globalization.CultureInfo]::CurrentUICulture.TwoLetterISOLanguageName -eq 'fr')
function T([string]$fr, [string]$en) { if ($Fr) { $fr } else { $en } }

# Resource IDs of the icons inside GoogleDriveFS.exe (checked against Drive for desktop 133.x).
# Every icon is verified by its dominant colour before use, so a Drive update that moves them is detected.
$Apps = @(
    @{ Key = 'Docs';   Def = 'DOCS';   Exe = 'OpenInDocs.exe';   Name = 'Google Docs';   Type = (T 'Document' 'Document');         Colour = 'blue';   FileIcon = 6;  AppIcon = 125; Exts = @('.docx', '.doc', '.docm', '.odt') },
    @{ Key = 'Sheets'; Def = 'SHEETS'; Exe = 'OpenInSheets.exe'; Name = 'Google Sheets'; Type = (T 'Classeur' 'Spreadsheet');      Colour = 'green';  FileIcon = 16; AppIcon = 126; Exts = @('.xlsx', '.xls', '.xlsm', '.ods', '.csv') },
    @{ Key = 'Slides'; Def = 'SLIDES'; Exe = 'OpenInSlides.exe'; Name = 'Google Slides'; Type = (T 'Présentation' 'Presentation'); Colour = 'yellow'; FileIcon = 17; AppIcon = 127; Exts = @('.pptx', '.ppt', '.pptm', '.ppsx', '.odp') }
)
$OldProgIds = @('OuvrirDansDocs.Fichier', 'OuvrirDansDocs.Docs', 'OuvrirDansDocs.Sheets', 'OuvrirDansDocs.Slides')
$OldExes    = @('OuvrirDocs.exe', 'OuvrirSheets.exe', 'OuvrirSlides.exe')
$AllExts    = @($Apps | ForEach-Object { $_.Exts }) + @('.xls', '.xlsx', '.ppt', '.pptx') | Select-Object -Unique

# ------------------------------------------------------------------------------------------------
# Reads the icons embedded in an .exe (PE resources RT_GROUP_ICON / RT_ICON) and rebuilds .ico files.
# ------------------------------------------------------------------------------------------------
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
public static class PeIcons
{
    static uint U32(byte[] d, int o) { return BitConverter.ToUInt32(d, o); }
    static int U16(byte[] d, int o) { return BitConverter.ToUInt16(d, o); }

    class Ent { public long Key; public uint Ptr; }

    public static Dictionary<int, byte[]> Extract(byte[] d, int[] wanted)
    {
        Dictionary<int, byte[]> result = new Dictionary<int, byte[]>();
        int pe = (int)U32(d, 0x3C);
        if (d[pe] != 'P' || d[pe + 1] != 'E') return result;
        int nsec = U16(d, pe + 6);
        int optsz = U16(d, pe + 20);
        int opt = pe + 24;
        int magic = U16(d, opt);
        int ddir = opt + (magic == 0x10b ? 96 : 112);
        uint rsrcRva = U32(d, ddir + 16);
        if (rsrcRva == 0) return result;

        uint[] va = new uint[nsec]; uint[] sz = new uint[nsec]; uint[] raw = new uint[nsec];
        int so = opt + optsz;
        for (int i = 0; i < nsec; i++)
        {
            uint vsize = U32(d, so + i * 40 + 8);
            va[i] = U32(d, so + i * 40 + 12);
            uint rawsize = U32(d, so + i * 40 + 16);
            raw[i] = U32(d, so + i * 40 + 20);
            sz[i] = Math.Max(vsize, rawsize);
        }
        Func<uint, int> r2o = delegate (uint rva)
        {
            for (int i = 0; i < nsec; i++)
                if (rva >= va[i] && rva < va[i] + sz[i]) return (int)(rva - va[i] + raw[i]);
            throw new Exception("RVA");
        };
        int rbase = r2o(rsrcRva);

        Func<int, List<Ent>> readDir = delegate (int off)
        {
            List<Ent> l = new List<Ent>();
            int named = U16(d, off + 12), ids = U16(d, off + 14);
            for (int i = 0; i < named + ids; i++)
            {
                uint nm = U32(d, off + 16 + i * 8);
                uint ptr = U32(d, off + 20 + i * 8);
                Ent e = new Ent();
                e.Key = ((nm & 0x80000000u) != 0) ? -1 : (long)nm;
                e.Ptr = ptr;
                l.Add(e);
            }
            return l;
        };
        Func<uint, byte[]> leaf = delegate (uint ptr)
        {
            while ((ptr & 0x80000000u) != 0) ptr = readDir(rbase + (int)(ptr & 0x7fffffffu))[0].Ptr;
            uint rva = U32(d, rbase + (int)ptr);
            uint size = U32(d, rbase + (int)ptr + 4);
            byte[] r = new byte[size];
            Array.Copy(d, r2o(rva), r, 0, (int)size);
            return r;
        };

        uint iconDir = 0, groupDir = 0;
        foreach (Ent e in readDir(rbase)) { if (e.Key == 3) iconDir = e.Ptr; else if (e.Key == 14) groupDir = e.Ptr; }
        if (iconDir == 0 || groupDir == 0) return result;

        Dictionary<long, uint> icons = new Dictionary<long, uint>();
        foreach (Ent e in readDir(rbase + (int)(iconDir & 0x7fffffffu))) if (e.Key >= 0) icons[e.Key] = e.Ptr;
        Dictionary<long, uint> groups = new Dictionary<long, uint>();
        foreach (Ent e in readDir(rbase + (int)(groupDir & 0x7fffffffu))) if (e.Key >= 0) groups[e.Key] = e.Ptr;

        foreach (int id in wanted)
        {
            uint gp;
            if (!groups.TryGetValue(id, out gp)) continue;
            byte[] g = leaf(gp);
            int cnt = U16(g, 4);
            List<byte[]> images = new List<byte[]>();
            List<byte[]> heads = new List<byte[]>();
            bool ok = true;
            for (int i = 0; i < cnt; i++)
            {
                int o = 6 + i * 14;
                int iid = U16(g, o + 12);
                uint ip;
                if (!icons.TryGetValue(iid, out ip)) { ok = false; break; }
                byte[] img = leaf(ip);
                byte[] head = new byte[16];
                Array.Copy(g, o, head, 0, 12);                       // width .. bytesInRes (12 bytes)
                Array.Copy(BitConverter.GetBytes((uint)img.Length), 0, head, 8, 4);
                images.Add(img); heads.Add(head);
            }
            if (!ok) continue;
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            ms.Write(new byte[] { 0, 0, 1, 0 }, 0, 4);
            ms.Write(BitConverter.GetBytes((ushort)cnt), 0, 2);
            int offset = 6 + 16 * cnt;
            for (int i = 0; i < cnt; i++)
            {
                Array.Copy(BitConverter.GetBytes((uint)offset), 0, heads[i], 12, 4);
                ms.Write(heads[i], 0, 16);
                offset += images[i].Length;
            }
            for (int i = 0; i < cnt; i++) ms.Write(images[i], 0, images[i].Length);
            result[id] = ms.ToArray();
        }
        return result;
    }
}
'@

Add-Type -AssemblyName System.Drawing

# Dominant colour (blue / green / yellow) of the 48x48 frame of an .ico: protects against a wrong icon.
function Get-IconColour([byte[]]$ico) {
    try {
        $n = [BitConverter]::ToUInt16($ico, 4)
        $bmp = $null
        for ($i = 0; $i -lt $n; $i++) {
            $o = 6 + 16 * $i
            if ($ico[$o] -ne 48) { continue }
            $size = [BitConverter]::ToUInt32($ico, $o + 8); $off = [BitConverter]::ToUInt32($ico, $o + 12)
            if ($ico[$off] -eq 0x89 -and $ico[$off + 1] -eq 0x50) {          # PNG frame
                $frame = New-Object byte[] $size; [Array]::Copy($ico, $off, $frame, 0, $size)
                $bmp = [System.Drawing.Bitmap]::FromStream((New-Object IO.MemoryStream(, $frame)))
            } else {                                                          # classic bitmap frame
                $bmp = (New-Object System.Drawing.Icon((New-Object IO.MemoryStream(, $ico)), 48, 48)).ToBitmap()
            }
            break
        }
        if (-not $bmp) { return $null }
        $c = @{ blue = 0; green = 0; yellow = 0 }
        for ($y = 0; $y -lt $bmp.Height; $y++) {
            for ($x = 0; $x -lt $bmp.Width; $x++) {
                $p = $bmp.GetPixel($x, $y)
                if ($p.A -lt 200) { continue }
                $s = $p.GetSaturation(); $l = $p.GetBrightness()
                if ($s -lt 0.4 -or $l -lt 0.25 -or $l -gt 0.85) { continue }
                $h = $p.GetHue()
                if ($h -ge 195 -and $h -le 260) { $c.blue++ } elseif ($h -ge 100 -and $h -le 170) { $c.green++ } elseif ($h -ge 30 -and $h -le 65) { $c.yellow++ }
            }
        }
        $best = ($c.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1)
        if ($best.Value -lt 20) { return $null }
        return $best.Key
    } catch { return $null }
}

function Find-DriveExe {
    $cands = @()
    foreach ($base in @("$env:ProgramFiles\Google\Drive File Stream", "${env:ProgramFiles(x86)}\Google\Drive File Stream")) {
        if (Test-Path -LiteralPath $base) {
            $cands += Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue |
                      Where-Object { $_.Name -match '^\d+(\.\d+)+$' } |
                      Sort-Object { [version]$_.Name } -Descending |
                      ForEach-Object { Join-Path $_.FullName 'GoogleDriveFS.exe' }
        }
    }
    $running = Get-Process GoogleDriveFS -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($running -and $running.Path) { $cands += $running.Path }
    return ($cands | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)
}

# Extracts the icons into $Dest\icons. Returns $true if all six icons were found and verified.
function Export-Icons {
    $exe = Find-DriveExe
    if (-not $exe) { Write-Host (T "  Google Drive pour ordinateur introuvable : icônes par défaut." "  Google Drive for desktop not found: default icons will be used.") -ForegroundColor Yellow; return $false }
    Write-Host (T "  Icônes lues dans : $exe" "  Reading icons from: $exe")
    $ids = @($Apps | ForEach-Object { $_.FileIcon; $_.AppIcon })
    $found = [PeIcons]::Extract([IO.File]::ReadAllBytes($exe), [int[]]$ids)
    $ok = $true
    foreach ($a in $Apps) {
        foreach ($pair in @(@('file', $a.FileIcon), @('app', $a.AppIcon))) {
            $id = $pair[1]
            $target = Join-Path $Dest ("icons\{0}-{1}.ico" -f $pair[0], $a.Key.ToLower())
            if ($found.ContainsKey($id) -and (Get-IconColour $found[$id]) -eq $a.Colour) {
                [IO.File]::WriteAllBytes($target, $found[$id])
            } else { $ok = $false }
        }
    }
    if (-not $ok) { Write-Host (T "  Les icônes n'ont pas pu être vérifiées (version de Drive différente ?) : icônes par défaut." "  Icons could not be verified (different Drive version?): default icons will be used.") -ForegroundColor Yellow }
    return $ok
}

# Creates a registry key only if it does not exist (New-Item -Force would empty an existing key).
function Ensure-Key([string]$p) { if (-not (Test-Path -LiteralPath $p)) { New-Item -Path $p -Force | Out-Null } }

function Clear-Registry {
    foreach ($e in $AllExts) {
        foreach ($id in ($OldProgIds + ($Apps | ForEach-Object { "OpenInDocs.$($_.Key)" }))) {
            Remove-ItemProperty -LiteralPath "$Classes\$e\OpenWithProgids" -Name $id -ErrorAction SilentlyContinue
        }
    }
    foreach ($id in ($OldProgIds + ($Apps | ForEach-Object { "OpenInDocs.$($_.Key)" }))) {
        Remove-Item -LiteralPath "$Classes\$id" -Recurse -Force -ErrorAction SilentlyContinue
    }
    foreach ($x in ($OldExes + ($Apps | ForEach-Object { $_.Exe }))) {
        Remove-Item -LiteralPath "$Classes\Applications\$x" -Recurse -Force -ErrorAction SilentlyContinue
    }
    foreach ($a in $Apps) { Remove-ItemProperty -LiteralPath $RegApps -Name $a.Name -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath 'HKCU:\Software\OuvrirDansDocs' -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath 'HKCU:\Software\OpenInDocs' -Recurse -Force -ErrorAction SilentlyContinue
}

function Notify-Shell {
    if (-not ('W.S' -as [type])) {
        Add-Type -Namespace W -Name S -MemberDefinition '[DllImport("shell32.dll")] public static extern void SHChangeNotify(int e, int f, IntPtr a, IntPtr b);'
    }
    [W.S]::SHChangeNotify(0x08000000, 0, [IntPtr]::Zero, [IntPtr]::Zero)
}

try {
    if ($Uninstall) {
        Clear-Registry
        Remove-Item -LiteralPath $Dest -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $OldDest -Recurse -Force -ErrorAction SilentlyContinue
        Notify-Shell
        Write-Host (T "Désinstallé. Au prochain double-clic, Windows demandera avec quelle application ouvrir le fichier." "Uninstalled. Next time you double-click a file, Windows will ask which app to use.")
        Read-Host (T "Entrée pour fermer" "Press Enter to close") | Out-Null
        return
    }

    Write-Host "OpenInDocs" -ForegroundColor Cyan

    # 1. C# compiler shipped with Windows (.NET Framework 4)
    $csc = @("$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
             "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $csc) { throw (T "Le compilateur .NET Framework (csc.exe) est introuvable." "The .NET Framework compiler (csc.exe) was not found.") }
    $cs = Join-Path $Root 'src\OpenInDocs.cs'
    if (-not (Test-Path -LiteralPath $cs)) { throw (T "src\OpenInDocs.cs est introuvable : lance l'installateur depuis le dossier du projet." "src\OpenInDocs.cs not found: run the installer from the project folder.") }

    # 2. Previous settings (OuvrirDansDocs v2 or an earlier OpenInDocs) are kept
    $oldIni = Join-Path $OldDest 'OuvrirDansDocs.ini'
    $keepIni = $null
    if (Test-Path -LiteralPath (Join-Path $Dest 'OpenInDocs.ini')) { $keepIni = Get-Content -LiteralPath (Join-Path $Dest 'OpenInDocs.ini') -Raw -Encoding UTF8 }
    elseif (Test-Path -LiteralPath $oldIni) { $keepIni = Get-Content -LiteralPath $oldIni -Raw -Encoding UTF8; Write-Host (T "  Réglages de l'ancienne version repris." "  Settings from the previous version kept.") }

    Clear-Registry
    Remove-Item -LiteralPath $OldDest -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Force -Path (Join-Path $Dest 'icons') | Out-Null

    $iniPath = Join-Path $Dest 'OpenInDocs.ini'
    if ($keepIni) { [IO.File]::WriteAllText($iniPath, $keepIni, (New-Object Text.UTF8Encoding($true))) }
    else {
        $tpl = Get-Content -LiteralPath (Join-Path $Root 'config\OpenInDocs.ini') -Raw -Encoding UTF8
        $tpl = $tpl.Replace('@TARGET@', (T 'Ouverts depuis le PC' 'Opened from PC'))
        [IO.File]::WriteAllText($iniPath, $tpl, (New-Object Text.UTF8Encoding($true)))
    }

    # 3. Icons taken from the user's own Drive for desktop installation
    $haveIcons = Export-Icons

    # 4. Build the 3 launchers (no console window)
    foreach ($a in $Apps) {
        $out = Join-Path $Dest $a.Exe
        $cscArgs = @('/nologo', '/target:winexe', '/optimize+', "/out:$out", "/define:$($a.Def)", '/r:System.Windows.Forms.dll', '/r:System.Drawing.dll')
        $ico = Join-Path $Dest ("icons\app-{0}.ico" -f $a.Key.ToLower())
        if ($haveIcons -and (Test-Path -LiteralPath $ico)) { $cscArgs += "/win32icon:$ico" }
        Write-Host (T "  Compilation de $($a.Exe)..." "  Building $($a.Exe)...")
        $ErrorActionPreference = 'Continue'
        $log = & $csc @cscArgs $cs 2>&1
        $code = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        if ($code -ne 0) { $log | ForEach-Object { Write-Host "    $_" -ForegroundColor Yellow }; throw (T "La compilation de $($a.Exe) a échoué (détails ci-dessus)." "Building $($a.Exe) failed (see details above).") }
    }

    # 5. Register with Windows
    foreach ($a in $Apps) {
        $progId = "OpenInDocs.$($a.Key)"
        $exe    = Join-Path $Dest $a.Exe
        $cmd    = "`"$exe`" `"%1`""
        $fileIco = Join-Path $Dest ("icons\file-{0}.ico" -f $a.Key.ToLower())

        $pk = "$Classes\$progId"
        Ensure-Key "$pk\shell\open\command"; Ensure-Key "$pk\Application"
        Set-Item         -LiteralPath $pk -Value "$($a.Type) ($($a.Name))"
        if ($haveIcons -and (Test-Path -LiteralPath $fileIco)) { Ensure-Key "$pk\DefaultIcon"; Set-Item -LiteralPath "$pk\DefaultIcon" -Value $fileIco }
        Set-Item         -LiteralPath "$pk\shell\open\command" -Value $cmd
        Set-ItemProperty -LiteralPath "$pk\shell\open" -Name 'FriendlyAppName' -Value $a.Name
        Set-ItemProperty -LiteralPath "$pk\Application" -Name 'ApplicationName' -Value $a.Name
        Set-ItemProperty -LiteralPath "$pk\Application" -Name 'ApplicationIcon' -Value "$exe,0"

        $ak = "$Classes\Applications\$($a.Exe)"
        Ensure-Key "$ak\shell\open\command"; Ensure-Key "$ak\SupportedTypes"
        Set-ItemProperty -LiteralPath $ak -Name 'FriendlyAppName' -Value $a.Name
        Set-Item         -LiteralPath "$ak\shell\open\command" -Value $cmd

        $ck = "HKCU:\Software\OpenInDocs\$($a.Key)\Capabilities"
        Ensure-Key "$ck\FileAssociations"
        Set-ItemProperty -LiteralPath $ck -Name 'ApplicationName' -Value $a.Name
        Set-ItemProperty -LiteralPath $ck -Name 'ApplicationDescription' -Value (T "Ouvre les fichiers dans $($a.Name) via Google Drive" "Opens files in $($a.Name) through Google Drive")
        Set-ItemProperty -LiteralPath $ck -Name 'ApplicationIcon' -Value "$exe,0"

        foreach ($e in $a.Exts) {
            Set-ItemProperty -LiteralPath "$ck\FileAssociations" -Name $e -Value $progId
            New-ItemProperty -Force -LiteralPath "$ak\SupportedTypes" -Name $e -PropertyType String -Value '' | Out-Null
            Ensure-Key "$Classes\$e\OpenWithProgids"
            New-ItemProperty -Force -LiteralPath "$Classes\$e\OpenWithProgids" -Name $progId -PropertyType String -Value '' | Out-Null
        }
        Ensure-Key $RegApps
        Set-ItemProperty -LiteralPath $RegApps -Name $a.Name -Value "Software\OpenInDocs\$($a.Key)\Capabilities"
    }
    Notify-Shell

    Write-Host ""
    Write-Host (T "Installé dans $Dest" "Installed in $Dest") -ForegroundColor Green
    Write-Host ""
    Write-Host (T "Dernière étape : choisir Google Docs / Sheets / Slides comme application par défaut." "Last step: choose Google Docs / Sheets / Slides as the default app.")
    Write-Host (T "La page « Applications par défaut » de Windows va s'ouvrir sur Google Docs : clique « Définir par défaut »." "Windows' \"Default apps\" page will open on Google Docs: click \"Set default\".")
    Write-Host (T "Fais de même pour Google Sheets et Google Slides." "Do the same for Google Sheets and Google Slides.")
    Write-Host (T "(Ou : double-clique un .docx > choisis « Google Docs » > Toujours.)" "(Or: double-click a .docx > pick \"Google Docs\" > Always.)")
    Start-Process 'ms-settings:defaultapps?registeredAppUser=Google%20Docs' -ErrorAction SilentlyContinue
}
catch {
    Write-Host ""
    Write-Host ((T "ERREUR : " "ERROR: ") + $_.Exception.Message) -ForegroundColor Red
    Write-Host (T "Ouvre une issue sur GitHub avec une capture de cette fenêtre." "Please open a GitHub issue with a screenshot of this window.")
}
Read-Host (T "Entrée pour fermer" "Press Enter to close") | Out-Null
