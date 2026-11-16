<#
  Folder Triage
  ============================================================
  Open a folder full of files/documents one at a time, look at each,
  press one key, and it files it away into a category folder and
  opens the next one.

  Put this script (and triage.cmd) in the folder that directly holds
  the items, then double-click triage.cmd.

  Keys:
    1 / 2 / 3 ...   assign to category 1 / 2 / 3 ...
    Enter           assign to category 1
    s               skip (leave in place, logged only)
    r               reopen current item
    o               open current item's folder
    k               force-close the viewer and reopen current item
    q               quit (progress is saved, resumes next run)
#>

[CmdletBinding()]
param(
    [string]$Path,
    [string[]]$Categories = @('好', '中', '坏'),
    [string]$Exts = 'pptx,ppt,pptm,potx,potm,doc,docx,pdf,odt,rtf,xls,xlsx,csv,odp,key,pages,jpg,jpeg,png,gif,bmp,webp,mp4,mov,mkv,mp3,wav,md,txt,html',
    [int]$Start = 0,
    [switch]$Any,
    [switch]$NoPreview,
    [switch]$Undo,
    [switch]$Portable,
    [switch]$PinConsole
)

$ErrorActionPreference = 'Stop'

try { chcp 65001 > $null } catch {}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$root = if ($Path) { (Resolve-Path -LiteralPath $Path).Path } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$statePath = Join-Path $root '_triage_state.json'
$logPath = Join-Path $root '_triage_log.csv'

$extList = @()
if (-not $Any) {
    $extList = @($Exts.Split(',') | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ } | ForEach-Object { if ($_.StartsWith('.')) { $_ } else { '.' + $_ } })
}

$sig = @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public class TriageUi {
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
  delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr SetActiveWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool c);
  [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr h, uint flags);
  [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
  [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();

  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }

  public static readonly IntPtr TOPMOST = new IntPtr(-1);
  public static readonly IntPtr NOTOPMOST = new IntPtr(-2);
  const uint SWP_NOSIZE = 0x0001, SWP_NOMOVE = 0x0002, SWP_SHOWWINDOW = 0x0040;

  public static List<string> Windows() {
    var res = new List<string>();
    EnumWindows((h, l) => {
      if (!IsWindowVisible(h)) return true;
      var t = new StringBuilder(1024); GetWindowText(h, t, 1024);
      if (t.Length == 0) return true;
      var c = new StringBuilder(256); GetClassName(h, c, 256);
      uint pid; GetWindowThreadProcessId(h, out pid);
      res.Add(h.ToString() + "|" + pid + "|" + c.ToString() + "|" + t.ToString().Replace("|", " "));
      return true;
    }, IntPtr.Zero);
    return res;
  }

  public static int PidOf(IntPtr h) {
    uint pid; GetWindowThreadProcessId(h, out pid); return (int)pid;
  }

  public static int WidthOf(IntPtr h) {
    RECT r; GetWindowRect(h, out r); return r.Right - r.Left;
  }

  public static void Front(IntPtr h) {
    if (h == IntPtr.Zero) return;
    uint pid; uint tid = GetWindowThreadProcessId(h, out pid);
    uint cur = GetCurrentThreadId();
    AttachThreadInput(cur, tid, true);
    ShowWindow(h, 9);
    BringWindowToTop(h);
    SetForegroundWindow(h);
    SetActiveWindow(h);
    AttachThreadInput(cur, tid, false);
  }

  public static void Place(IntPtr h, int x, int y, int w, int hgt, bool topmost) {
    if (h == IntPtr.Zero) return;
    ShowWindow(h, 9);
    SetWindowPos(h, topmost ? TOPMOST : NOTOPMOST, x, y, w, hgt, SWP_SHOWWINDOW);
  }

  public static bool Pin(IntPtr h, int x, int y) {
    if (h == IntPtr.Zero) return false;
    return SetWindowPos(h, TOPMOST, x, y, 0, 0, SWP_NOSIZE | SWP_SHOWWINDOW);
  }
}
'@
Add-Type -TypeDefinition $sig -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Windows.Forms

$script:consoleOk = $PinConsole -and (-not [Console]::IsInputRedirected) -and (-not [Console]::IsOutputRedirected)


# ---------------------------------------------------------------- window helpers

function Get-Screen { [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea }

function Get-AncestorPids {
    $list = New-Object System.Collections.ArrayList
    $cur = $PID
    for ($i = 0; $i -lt 6; $i++) {
        try {
            $p = Get-CimInstance Win32_Process -Filter "ProcessId=$cur" -ErrorAction Stop
            if (-not $p) { break }
            $parent = [int]$p.ParentProcessId
            if ($parent -le 4 -or $list -contains $parent) { break }
            [void]$list.Add($parent)
            $cur = $parent
        } catch { break }
    }
    return $list
}

function Get-ConsoleHwnd {
    if (-not $script:consoleOk) { return [IntPtr]::Zero }
    $h = [TriageUi]::GetConsoleWindow()
    if ($h -eq [IntPtr]::Zero) { return [IntPtr]::Zero }
    $root2 = [TriageUi]::GetAncestor($h, 2)
    if ($root2 -ne [IntPtr]::Zero) { return $root2 }
    return $h
}

function Set-ConsoleLayout {
    $h = Get-ConsoleHwnd
    if ($h -eq [IntPtr]::Zero) { return }
    try {
        if ([Console]::BufferWidth -lt 100) { [Console]::BufferWidth = 100 }
        [Console]::WindowWidth = 96
        [Console]::WindowHeight = 22
    } catch { }
    $wa = Get-Screen
    [TriageUi]::Pin($h, $wa.Left + 12, $wa.Top + 12) | Out-Null
}

function Focus-Console {
    $h = Get-ConsoleHwnd
    if ($h -eq [IntPtr]::Zero) { return }
    [TriageUi]::Front($h)
    $wa = Get-Screen
    [TriageUi]::Pin($h, $wa.Left + 12, $wa.Top + 12) | Out-Null
}

function Get-WindowMap {
    $map = @{}
    foreach ($w in [TriageUi]::Windows()) {
        $p = $w.Split('|')
        if ($p.Length -lt 4) { continue }
        $map[[IntPtr][long]$p[0]] = @{ Pid = [int]$p[1]; Cls = $p[2]; Title = $p[3] }
    }
    return $map
}

function Place-Viewer([IntPtr]$h) {
    $wa = Get-Screen
    $cw = 0
    $ch = Get-ConsoleHwnd
    if ($ch -ne [IntPtr]::Zero) { $cw = [TriageUi]::WidthOf($ch) }
    if ($cw -le 200) { $cw = 780 }
    $x = $wa.Left + $cw + 28
    $w = $wa.Right - $x - 16
    if ($w -lt 640) {
        $x = $wa.Left + [int]($wa.Width * 0.42)
        $w = $wa.Right - $x - 16
    }
    $hgt = $wa.Height - 48
    [TriageUi]::Place($h, $x, $wa.Top + 12, $w, $hgt, $false)
}

# ---------------------------------------------------------------- open / close

function Convert-Template([string]$src, [string]$dst) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $dst) { Remove-Item -LiteralPath $dst -Force }
    $zin = [System.IO.Compression.ZipFile]::OpenRead($src)
    try {
        $zout = [System.IO.Compression.ZipFile]::Open($dst, 'Create')
        try {
            foreach ($e in $zin.Entries) {
                $ne = $zout.CreateEntry($e.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
                $os = $ne.Open()
                try {
                    $is = $e.Open()
                    try {
                        if ($e.FullName -eq '[Content_Types].xml') {
                            $sr = New-Object System.IO.StreamReader($is)
                            $xml = $sr.ReadToEnd()
                            $sr.Close()
                            $xml = $xml.Replace('presentationml.template.main+xml', 'presentationml.presentation.main+xml')
                            $xml = $xml.Replace('ms-powerpoint.template.macroEnabled.main+xml', 'ms-powerpoint.presentation.macroEnabled.main+xml')
                            $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($xml)
                            $os.Write($bytes, 0, $bytes.Length)
                        } else {
                            $is.CopyTo($os)
                        }
                    } finally { $is.Close() }
                } finally { $os.Close() }
            }
        } finally { $zout.Dispose() }
    } finally { $zin.Dispose() }
}

function Open-Item([string]$file) {
    $before = Get-WindowMap
    $consoleHwnd = Get-ConsoleHwnd

    $target = $file
    $ext = [System.IO.Path]::GetExtension($file).ToLowerInvariant()
    if ($ext -eq '.potx' -or $ext -eq '.potm' -or $ext -eq '.pot') {
        $tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) 'triage_tmp'
        if (-not (Test-Path -LiteralPath $tmpDir)) { New-Item -ItemType Directory -Path $tmpDir | Out-Null }
        $newExt = if ($ext -eq '.potm') { '.pptm' } elseif ($ext -eq '.pot') { '.ppt' } else { '.pptx' }
        $candidate = Join-Path $tmpDir ([System.IO.Path]::GetFileNameWithoutExtension($file) + $newExt)
        if ($ext -ne '.pot') {
            try { Convert-Template $file $candidate; $target = $candidate } catch { $target = $file }
        }
    }

    Write-Host "        打开中…" -ForegroundColor DarkGray
    Start-Process -FilePath $target

    $deadline = (Get-Date).AddSeconds(45)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 600
        $now = Get-WindowMap
        foreach ($k in $now.Keys) {
            if ($k -eq $consoleHwnd) { continue }
            if ($before.ContainsKey($k)) { continue }
            $info = $now[$k]
            if ($info.Pid -eq $PID) { continue }
            if ($info.Cls -in @('#32770', 'NUIDialog', 'tooltips_class32', 'TaskListThumbnailWnd')) { continue }
            Place-Viewer $k
            [TriageUi]::Front($k)
            Start-Sleep -Milliseconds 350
            Focus-Console
            return $k
        }
    }
    Focus-Console
    return [IntPtr]::Zero
}

function Close-Item([IntPtr]$h) {
    if ($h -ne [IntPtr]::Zero -and [TriageUi]::IsWindow($h)) {
        for ($i = 0; $i -lt 12; $i++) {
            [TriageUi]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            Start-Sleep -Milliseconds 500
            if (-not [TriageUi]::IsWindow($h)) { return }
        }
        $pid2 = [TriageUi]::PidOf($h)
        if ($pid2 -gt 0) { try { (Get-Process -Id $pid2 -ErrorAction Stop).Kill() } catch { } }
        Start-Sleep -Milliseconds 800
    }
}

function Close-AllViewers {
    foreach ($h in $script:openedHandles) { Close-Item $h }
    $script:openedHandles = @()
    Close-StrayDialogs
    Focus-Console
}

function Reset-Viewers {
    foreach ($h in $script:openedHandles) { Close-Item $h }
    $script:openedHandles = @()
    foreach ($q in $script:openedPids) {
        try { Stop-Process -Id $q -Force -ErrorAction Stop } catch { }
    }
    $script:openedPids = @()
    Start-Sleep -Milliseconds 700
    Focus-Console
}

function Wait-FileReleased([string]$file, [int]$TimeoutSec = 10) {
    if (-not $file) { return $true }
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        try {
            $fs = [System.IO.File]::Open($file, 'Open', 'Read', 'None')
            $fs.Close()
            return $true
        } catch { Start-Sleep -Milliseconds 350 }
    }
    return $false
}

function Ensure-Released([string]$file) {
    if (Wait-FileReleased $file 10) { return $true }
    foreach ($q in $script:openedPids) {
        try { Stop-Process -Id $q -Force -ErrorAction Stop } catch { }
    }
    return (Wait-FileReleased $file 8)
}

function Close-StrayDialogs {
    $consoleHwnd = Get-ConsoleHwnd
    foreach ($w in [TriageUi]::Windows()) {
        $p = $w.Split('|')
        if ($p.Length -lt 4) { continue }
        $h = [IntPtr][long]$p[0]
        if ($h -eq $consoleHwnd) { continue }
        if ($p[2] -in @('#32770', 'NUIDialog')) {
            [TriageUi]::PostMessage($h, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
        }
    }
}

$script:openedHandles = @()

# ---------------------------------------------------------------- scan / state

function Test-Wanted([System.IO.FileInfo]$f) {
    if ($Any) { return $true }
    return ($extList -contains $f.Extension.ToLowerInvariant())
}

function Get-LoosePpt([string]$dir) {
    Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue |
        Where-Object { Test-Wanted $_ } |
        Sort-Object Length -Descending |
        Select-Object -First 1
}

function Load-State {
    if (Test-Path -LiteralPath $statePath) {
        try { return (Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { }
    }
    return $null
}

function Save-State($s) { $s | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding UTF8 }

function Sync-State {
    $old = Load-State
    $prevCurrent = 1
    $items = @()
    if ($old) {
        if ($old.PSObject.Properties['Current'] -and [int]$old.Current -gt 0) { $prevCurrent = [int]$old.Current }
        foreach ($it in @($old.Items)) {
            $items += [pscustomobject]@{
                Kind    = if ($it.PSObject.Properties['Kind'] -and $it.Kind) { $it.Kind } else { 'dir' }
                Name    = $it.Name
                File    = $it.File
                Verdict = if ($it.PSObject.Properties['Verdict']) { $it.Verdict } else { $null }
                Time    = if ($it.PSObject.Properties['Time']) { $it.Time } else { $null }
            }
        }
    }
    $state = [pscustomobject]@{ Current = $prevCurrent; Items = @($items) }
    if ($Start -gt 0) { $state.Current = $Start }

    $known = @{}
    foreach ($it in $state.Items) { $known[$it.Name] = 1 }

    Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '_*' -and $_.Name -notin $Categories } |
        Sort-Object Name |
        ForEach-Object {
            if (-not $known.ContainsKey($_.Name)) {
                $f = Get-LoosePpt $_.FullName
                if ($f) {
                    $state.Items += [pscustomobject]@{ Kind = 'dir'; Name = $_.Name; File = $f.FullName; Verdict = $null; Time = $null }
                    $known[$_.Name] = 1
                }
            }
        }

    Get-ChildItem -LiteralPath $root -File -ErrorAction SilentlyContinue |
        Where-Object { Test-Wanted $_ -and -not $_.Name.StartsWith('_') } |
        Sort-Object Name |
        ForEach-Object {
            $n = $_.Name
            if (-not $known.ContainsKey($n)) {
                $state.Items += [pscustomobject]@{ Kind = 'file'; Name = $n; File = $_.FullName; Verdict = $null; Time = $null }
                $known[$n] = 1
            }
        }

    $alive = @()
    foreach ($it in @($state.Items)) {
        if ($it.Verdict) { $alive += $it; continue }
        if ($it.Kind -eq 'file') {
            if ($it.File -and (Test-Path -LiteralPath $it.File)) { $alive += $it }
        } else {
            $d = Join-Path $root $it.Name
            if (Test-Path -LiteralPath $d) {
                if (-not $it.File -or -not (Test-Path -LiteralPath $it.File)) {
                    $f = Get-LoosePpt $d
                    if ($f) { $it.File = $f.FullName }
                }
                if ($it.File) { $alive += $it }
            }
        }
    }
    $state.Items = @($alive)
    Save-State $state
    return $state
}

function Count-Bucket([string]$b) {
    $p = Join-Path $root $b
    if (-not (Test-Path -LiteralPath $p)) { return 0 }
    $n = @(Get-ChildItem -LiteralPath $p -Directory -ErrorAction SilentlyContinue).Count
    if ($Any) {
        $n += @(Get-ChildItem -LiteralPath $p -File -ErrorAction SilentlyContinue).Count
    } else {
        $n += @(Get-ChildItem -LiteralPath $p -File -ErrorAction SilentlyContinue | Where-Object { $extList -contains $_.Extension.ToLowerInvariant() }).Count
    }
    return $n
}

function Write-Log($item, [string]$verdict, [string]$note) {
    $row = [pscustomobject]@{
        时间   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        名称   = $item.Name
        文件   = if ($item.File) { Split-Path $item.File -Leaf } else { '' }
        分类   = $verdict
        备注   = $note
    }
    if (Test-Path -LiteralPath $logPath) {
        $row | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8 -Append
    } else {
        $row | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
    }
}

function Move-Raw([string]$src, [string]$dst) {
    for ($i = 0; $i -lt 25; $i++) {
        try {
            if ([System.IO.Directory]::Exists($src)) {
                [System.IO.Directory]::Move($src, $dst)
                return $true
            }
            if ([System.IO.File]::Exists($src)) {
                [System.IO.File]::Move($src, $dst)
                return $true
            }
            $script:lastMoveError = 'source vanished'
            return $false
        } catch {
            $script:lastMoveError = $_.Exception.Message
            Start-Sleep -Milliseconds 700
        }
    }
    return $false
}

function Remove-EmptyJunk([string]$path) {
    try {
        if ([System.IO.Directory]::Exists($path) -and (Get-ChildItem -LiteralPath $path -Force -ErrorAction SilentlyContinue).Count -eq 0) {
            [System.IO.Directory]::Delete($path)
        }
    } catch { }
}

function Move-ToCategory($item, [string]$cat) {
    $dstDir = Join-Path $root $cat
    if (-not (Test-Path -LiteralPath $dstDir)) { New-Item -ItemType Directory -Path $dstDir | Out-Null }

    if ($item.Kind -eq 'file') {
        $src = $item.File
        if (-not (Test-Path -LiteralPath $src)) { return 'file not found' }
        $leaf = Split-Path $src -Leaf
        $final = Join-Path $dstDir $leaf
        if (Test-Path -LiteralPath $final) {
            $stem = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
            $ext = [System.IO.Path]::GetExtension($leaf)
            $n = 2
            while (Test-Path -LiteralPath (Join-Path $dstDir ("{0} ({1}){2}" -f $stem, $n, $ext))) { $n++ }
            $final = Join-Path $dstDir ("{0} ({1}){2}" -f $stem, $n, $ext)
        }
        if (Move-Raw $src $final) { return "moved -> $cat" }
        Reset-Viewers
        if (Move-Raw $src $final) { return "moved -> $cat" }
        return "MOVE FAILED: $script:lastMoveError"
    }

    $src = Join-Path $root $item.Name
    if (-not (Test-Path -LiteralPath $src)) { return 'folder not found' }
    $final = Join-Path $dstDir $item.Name
    if (Test-Path -LiteralPath $final) {
        Remove-EmptyJunk $final
        if (Test-Path -LiteralPath $final) {
            $n = 2
            while (Test-Path -LiteralPath ($final + " ($n)")) { $n++ }
            $final = $final + " ($n)"
        }
    }
    if (Move-Raw $src $final) { return "moved -> $cat" }
    Reset-Viewers
    if (Move-Raw $src $final) { return "moved -> $cat" }
    return "MOVE FAILED: $script:lastMoveError"
}
function Invoke-Undo {
    if (-not (Test-Path -LiteralPath $logPath)) { Write-Host "没有日志，无法撤销。" -ForegroundColor Yellow; return }
    $rows = @(Import-Csv -LiteralPath $logPath -Encoding UTF8)
    $done = @($rows | Where-Object { $Categories -contains $_.分类 })
    if ($done.Count -eq 0) { Write-Host "没有可撤销的移动。" ; return }
    Write-Host "将把 $($done.Count) 项移回原位（最近的在最前）。" -ForegroundColor Yellow
    $ans = Read-Host "确认撤销？(y/N)"
    if ($null -eq $ans) { return }
    if ($ans.Trim() -notmatch '^(y|yes|是)$') { return }
    $ok = 0
    foreach ($r in $done) {
        $src = Join-Path (Join-Path $root $r.分类) $r.名称
        $dst = Join-Path $root $r.名称
        if (Test-Path -LiteralPath $src) {
            try { Move-Item -LiteralPath $src -Destination $dst -Force; $ok++ } catch { Write-Host "失败：$($r.名称)" -ForegroundColor Yellow }
        }
    }
    "已撤销 $ok 项。"
}

# ---------------------------------------------------------------- preview

function Show-Counter($state) {
    $c = @{}
    foreach ($cat in $Categories) { $c[$cat] = 0 }
    foreach ($it in @($state.Items)) { if ($it.Verdict -and $c.ContainsKey($it.Verdict)) { $c[$it.Verdict]++ } }
    $a = @($Categories | ForEach-Object { "$_ $($c[$_])" }) -join '  '
    $b = @($Categories | ForEach-Object { "$_ $(Count-Bucket $_)" }) -join '  '
    return "   本次：$a      累计：$b"
}

function Show-Preview($state) {
    $pool = @($state.Items) | Where-Object { $_.File -and (Test-Path -LiteralPath $_.File) }
    if ($pool.Count -eq 0) { return }
    $n = [int][Math]::Round($pool.Count / 3)
    if ($n -lt 3) { $n = [Math]::Min(3, $pool.Count) }
    if ($n -gt $pool.Count) { $n = $pool.Count }
    $sample = @($pool | Get-Random -Count $n)

    Write-Host ""
    Write-Host "=== 随机预览 $n 份（共 $($pool.Count) 份） ===" -ForegroundColor Cyan
    Write-Host "只看不打分，看完按回车换下一份；预览里打 q 提前结束。" -ForegroundColor DarkGray
    Write-Host ""

    $k = 0
    foreach ($it in $sample) {
        $k++
        Write-Host ("-" * 72) -ForegroundColor DarkGray
        Write-Host "预览 [$k/$n] $($it.Name)" -ForegroundColor Cyan
        $h = Open-Item $it.File
        if ($h -ne [IntPtr]::Zero) { $script:openedHandles += $h }
        Focus-Console
        $a = Read-Host "        看完按回车换下一份（q 结束预览）"
        Close-AllViewers
        if ($null -ne $a -and $a.Trim() -match '^(q|quit|退出)$') { break }
    }
    Write-Host ""
    Write-Host "预览结束，开始正式分拣。" -ForegroundColor Green
    Write-Host ""
}

# ---------------------------------------------------------------- main

if ($Undo) { Invoke-Undo; return }

Set-ConsoleLayout
$state = Sync-State
$total = @($state.Items).Count

Write-Host ""
Write-Host "=== Folder Triage ===" -ForegroundColor Cyan
Write-Host "目录：$root"
Write-Host "待处理 $total 项，从第 $($state.Current) 项开始"
Write-Host ("分类：" + ($Categories -join ' / '))
Write-Host (Show-Counter $state)
$keyHelp = '1/2/3=分类  Enter=第一类  s=跳过  r=重开  o=文件夹  k=重置  q=退出'
Write-Host "按键：$keyHelp" -ForegroundColor DarkGray
Write-Host ""

if ($total -eq 0) {
    Write-Host "这个目录里没有待处理项。把脚本放进「直接装着各个条目」的目录再运行。" -ForegroundColor Yellow
    return
}

if (-not $NoPreview) {
    Focus-Console
    $a = Read-Host "开始前要不要先随机预览 1/3 找找感觉？(y=要 / 回车=直接开始)"
    if ($null -eq $a) { return }
    if ($a.Trim() -match '^(y|yes|是|要|好)$') { Show-Preview $state }
}

while ($true) {
    $idx = [int]$state.Current
    if ($idx -lt 1) { $idx = 1 }
    $item = @($state.Items)[$idx - 1]
    if (-not $item) { Write-Host "全部完成。" -ForegroundColor Green; break }
    if ($item.Verdict) { $state.Current = $idx + 1; Save-State $state; continue }

    if (-not $item.File -or -not (Test-Path -LiteralPath $item.File)) {
        Write-Host "[$idx/$total] $($item.Name)  -> 没有可打开的文件，自动跳过" -ForegroundColor Yellow
        $item.Verdict = '跳过'
        $item.Time = (Get-Date).ToString('s')
        Write-Log $item '跳过' 'no file'
        $state.Current = $idx + 1
        Save-State $state
        continue
    }

    $sizeMb = [Math]::Round((Get-Item -LiteralPath $item.File).Length / 1MB, 1)
    Write-Host ("-" * 72) -ForegroundColor DarkGray
    Write-Host "[$idx/$total] $($item.Name)" -ForegroundColor Cyan
    Write-Host "        文件：$(Split-Path $item.File -Leaf)   ${sizeMb} MB"
    Write-Host (Show-Counter $state)

    $handle = Open-Item $item.File
    if ($handle -ne [IntPtr]::Zero) {
        $script:openedHandles += $handle
        $script:openedPids += [TriageUi]::PidOf($handle)
    } else {
        Write-Host "        !! 没检测到新窗口，可能已经用别的方式打开了" -ForegroundColor Yellow
    }

    $verdict = $null
    while (-not $verdict) {
        Focus-Console
        $ans = Read-Host "        分类"
        if ($null -eq $ans) { $verdict = '__quit__'; break }
        $ans = $ans.Trim()
        if ($ans -match '^(q|quit|exit|退出)$') { $verdict = '__quit__'; break }
        if ($ans -match '^(s|skip|跳过)$') { $verdict = '跳过'; break }
        if ($ans -eq '') { $verdict = $Categories[0]; break }
        if ($ans -match '^(r|reopen|重开)$') { $verdict = '__reopen__'; break }
        if ($ans -match '^(k|reset|重置)$') { $verdict = '__reset__'; break }
        if ($ans -match '^(o|open|打开)$') {
            Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"{0}"' -f $item.File)
            $verdict = $null
            continue
        }
        $num = 0
        if ([int]::TryParse($ans, [ref]$num) -and $num -ge 1 -and $num -le $Categories.Count) {
            $verdict = $Categories[$num - 1]
            break
        }
        $hit = $Categories | Where-Object { $ans -eq $_ }
        if ($hit) { $verdict = @($hit)[0]; break }
        Write-Host "        可用：$($Categories -join ' / ')  或 1..$($Categories.Count)" -ForegroundColor Yellow
    }

    if ($verdict -eq '__quit__') {
        Close-AllViewers
        Save-State $state
        Write-Host ""
        Write-Host "已保存进度，下次从第 $($state.Current) 项继续。" -ForegroundColor Green
        break
    }

    if ($verdict -eq '__reopen__') { Close-AllViewers; continue }

    if ($verdict -eq '__reset__') { Reset-Viewers; continue }

    Close-AllViewers

    $note = if ($Categories -contains $verdict) { Move-ToCategory $item $verdict } else { 'skipped' }
    if ($note -like 'MOVE FAILED*') {
        Write-Host "        移动失败：$note" -ForegroundColor Red
        Write-Host "        已尝试自动清理，请再选一次这个分类。" -ForegroundColor Red
        continue
    }
    $item.Verdict = $verdict
    $item.Time = (Get-Date).ToString('s')
    Write-Log $item $verdict $note
    $state.Current = $idx + 1
    Save-State $state

    Write-Host "        => $verdict   ($note)" -ForegroundColor Green
    Write-Host (Show-Counter $state)
    if ($state.Current -gt $total) {
        Write-Host ""
        Write-Host "全部 $total 项处理完成。" -ForegroundColor Green
        break
    }
}

Write-Host ""
Write-Host "结果："
foreach ($c in $Categories) { Write-Host ("  {0}\{1} 项" -f $c, (Count-Bucket $c)) }
Write-Host "日志：_triage_log.csv   进度：_triage_state.json"
