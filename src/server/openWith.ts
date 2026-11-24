import { execFile, spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import { promisify } from 'node:util';
import { resolveInputPath } from './files.js';

const execFileAsync = promisify(execFile);

export interface OpenWithApp {
  id: string;
  name: string;
  exePath: string;
  isDefault: boolean;
  iconDataUri?: string;
}

export interface OpenWithResult {
  path: string;
  ext: string;
  platform: NodeJS.Platform;
  defaultApp: OpenWithApp | null;
  apps: OpenWithApp[];
}

interface CacheEntry {
  expiresAt: number;
  apps: OpenWithApp[];
  defaultApp: OpenWithApp | null;
}

const CACHE_TTL_MS = 5 * 60 * 1000;
const extCache = new Map<string, CacheEntry>();
const allowedExePaths = new Set<string>();

function normalizeExeKey(exePath: string): string {
  return path.normalize(exePath).toLowerCase();
}

function buildWin32DiscoveryScript(sanitizedExt: string): string {
  const escapedExt = sanitizedExt.replace(/'/g, "''");
  return `
$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Drawing

$HKCU = [Microsoft.Win32.Registry]::CurrentUser
$HKCR = [Microsoft.Win32.Registry]::ClassesRoot
$HKLM = [Microsoft.Win32.Registry]::LocalMachine

function RegVal($root, $sub, $name = '') {
  $k = $null
  try { $k = $root.OpenSubKey($sub); if ($k) { return $k.GetValue($name) } } catch {} finally { if ($k) { $k.Dispose() } }
  return $null
}

function RegHas($root, $sub, $name) {
  $k = $null
  try { $k = $root.OpenSubKey($sub); if ($k) { return ($k.GetValueNames() -contains $name) } } catch {} finally { if ($k) { $k.Dispose() } }
  return $false
}

function Resolve-Cmd($cmd) {
  if (-not $cmd -or -not ($cmd -is [string])) { return $null }
  $cmd = $cmd.Trim()
  $exe = $null
  if ($cmd.StartsWith('"')) {
    $end = $cmd.IndexOf('"', 1)
    if ($end -gt 1) { $exe = $cmd.Substring(1, $end - 1) }
  } else {
    $m = [regex]::Match($cmd, '^(.+?\\.exe)\\b', 'IgnoreCase')
    if ($m.Success) { $exe = $m.Groups[1].Value.Trim() }
  }
  if ($exe) { $exe = [Environment]::ExpandEnvironmentVariables($exe) }
  if ($exe -and [System.IO.File]::Exists($exe)) { return $exe }
  return $null
}

$ext = '${escapedExt}'
$seenExe = @{}
$seenName = @{}
$candidates = New-Object System.Collections.Generic.List[object]

function Push-Candidate($exe, $isDef, $fallbackName, $tier) {
  if (-not $exe) { return }
  $leaf = [System.IO.Path]::GetFileName($exe).ToLowerInvariant()
  if ($leaf -eq 'rundll32.exe' -or $leaf -eq 'cmd.exe' -or $leaf -eq 'openwith.exe') { return }
  $key = $exe.ToLowerInvariant()
  if ($seenExe.ContainsKey($key)) { return }

  $name = $fallbackName
  if ($name -and $name.StartsWith('@')) { $name = $null }
  try {
    $vi = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
    if ($vi.FileDescription -and $vi.FileDescription.Trim() -and $vi.FileDescription.Trim() -ne 'Application') {
      $name = $vi.FileDescription.Trim()
    } elseif ($vi.ProductName -and $vi.ProductName.Trim() -and $vi.ProductName.Trim() -ne 'Operating System') {
      $name = $vi.ProductName.Trim()
    }
  } catch {}
  if (-not $name) { $name = [System.IO.Path]::GetFileNameWithoutExtension($exe) }
  $nameKey = $name.Trim().ToLowerInvariant()
  if ($seenName.ContainsKey($nameKey)) { return }

  $seenExe[$key] = $true
  $seenName[$nameKey] = $true

  $iconDataUri = ''
  try {
    $ico = [System.Drawing.Icon]::ExtractAssociatedIcon($exe)
    if ($ico) {
      $bmp = New-Object System.Drawing.Bitmap(20, 20)
      $g = [System.Drawing.Graphics]::FromImage($bmp)
      $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
      $g.DrawIcon($ico, (New-Object System.Drawing.Rectangle(0, 0, 20, 20)))
      $ms = New-Object System.IO.MemoryStream
      $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
      $iconDataUri = 'data:image/png;base64,' + [Convert]::ToBase64String($ms.ToArray())
      $ms.Dispose(); $g.Dispose(); $bmp.Dispose(); $ico.Dispose()
    }
  } catch {}

  $candidates.Add([pscustomobject]@{
    id = $key
    name = $name
    exePath = $exe
    isDefault = [bool]$isDef
    tier = [int]$tier
    iconDataUri = $iconDataUri
  })
}

if ($ext) {
  $progIds = New-Object System.Collections.Generic.List[string]
  $uc = RegVal $HKCU "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\$ext\\UserChoice" 'ProgId'
  if ($uc) { $progIds.Add([string]$uc) }
  $def = RegVal $HKCR $ext ''
  if ($def -and -not $progIds.Contains([string]$def)) { $progIds.Add([string]$def) }

  foreach ($pair in @(
    @($HKCU, "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\$ext\\OpenWithProgids"),
    @($HKCR, "$ext\\OpenWithProgids")
  )) {
    $k = $pair[0].OpenSubKey($pair[1])
    if ($k) {
      foreach ($vn in $k.GetValueNames()) { if ($vn -and -not $progIds.Contains($vn)) { $progIds.Add($vn) } }
      $k.Dispose()
    }
  }

  $defaultAssigned = $false
  foreach ($p in $progIds) {
    $cmd = RegVal $HKCR "$p\\shell\\open\\command" ''
    $exe = Resolve-Cmd $cmd
    if (-not $exe) { continue }
    $fname = RegVal $HKCR "$p\\Application" 'ApplicationName'
    if (-not $fname) { $fname = RegVal $HKCR $p '' }
    $isDef = -not $defaultAssigned
    if ($isDef) { $defaultAssigned = $true }
    Push-Candidate $exe $isDef $fname 1
  }

  foreach ($pair in @(
    @($HKCU, "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\$ext\\OpenWithList"),
    @($HKCR, "$ext\\OpenWithList")
  )) {
    $k = $pair[0].OpenSubKey($pair[1])
    if ($k) {
      foreach ($vn in $k.GetValueNames()) {
        if ($vn -and $vn -ne 'MRUList') {
          $en = [string]$k.GetValue($vn)
          if ($en.EndsWith('.exe', 'OrdinalIgnoreCase')) {
            $exe = Resolve-Cmd (RegVal $HKCR "Applications\\$en\\shell\\open\\command" '')
            if (-not $exe) {
              foreach ($r in @($HKCU, $HKLM)) {
                $ap = RegVal $r "Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\$en" ''
                if ($ap -is [string]) {
                  $ap = [Environment]::ExpandEnvironmentVariables($ap.Trim('"'))
                  if ([System.IO.File]::Exists($ap)) { $exe = $ap; break }
                }
              }
            }
            Push-Candidate $exe $false (RegVal $HKCR "Applications\\$en" 'FriendlyAppName') 2
          }
        }
      }
      $k.Dispose()
    }
  }
}

$appKey = $HKCR.OpenSubKey('Applications')
if ($appKey) {
  foreach ($sk in $appKey.GetSubKeyNames()) {
    if (-not $sk.EndsWith('.exe', 'OrdinalIgnoreCase')) { continue }
    if (RegHas $HKCR "Applications\\$sk" 'NoOpenWith') { continue }
    $cmd = RegVal $HKCR "Applications\\$sk\\shell\\open\\command" ''
    $exe = Resolve-Cmd $cmd
    if (-not $exe) { continue }
    $tier = 3
    $supKey = $HKCR.OpenSubKey("Applications\\$sk\\SupportedTypes")
    if ($supKey) {
      $vals = $supKey.GetValueNames()
      $supKey.Dispose()
      if ($ext -and ($vals -contains $ext -or $vals -contains '*')) { $tier = 2 } else { continue }
    }
    Push-Candidate $exe $false (RegVal $HKCR "Applications\\$sk" 'FriendlyAppName') $tier
  }
  $appKey.Dispose()
}

$sorted = @($candidates | Sort-Object @{Expression='isDefault';Descending=$true}, @{Expression='tier';Ascending=$true}, @{Expression='name';Ascending=$true} | Select-Object -First 18)
ConvertTo-Json -InputObject @{ apps = $sorted } -Compress -Depth 4
`;
}

async function discoverWin32Apps(ext: string): Promise<OpenWithApp[]> {
  const sanitizedExt = /^\.[a-z0-9_-]{1,16}$/i.test(ext) ? ext.toLowerCase() : '';
  const script = buildWin32DiscoveryScript(sanitizedExt);
  const encoded = Buffer.from(script, 'utf16le').toString('base64');

  const { stdout } = await execFileAsync(
    'powershell.exe',
    ['-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
    {
      timeout: 8000,
      windowsHide: true,
      maxBuffer: 8 * 1024 * 1024,
      encoding: 'utf8',
    },
  );

  const raw = (stdout || '').trim();
  if (!raw) return [];
  const parsed = JSON.parse(raw) as { apps?: Array<OpenWithApp & { tier?: number }> };
  const list = Array.isArray(parsed?.apps) ? parsed.apps : [];
  return list
    .filter((item) => item && typeof item.exePath === 'string' && item.exePath.length > 0)
    .map((item) => {
      allowedExePaths.add(normalizeExeKey(item.exePath));
      return {
        id: item.id || normalizeExeKey(item.exePath),
        name: item.name || path.basename(item.exePath, '.exe'),
        exePath: item.exePath,
        isDefault: Boolean(item.isDefault),
        iconDataUri: item.iconDataUri || undefined,
      };
    });
}

export async function listOpenWithApps(rawPath: string): Promise<OpenWithResult> {
  const resolved = resolveInputPath(rawPath);
  const stat = await fs.stat(resolved);
  if (!stat.isFile()) {
    throw new Error(`Not a file: ${resolved}`);
  }

  const ext = path.extname(resolved).toLowerCase();
  const cacheKey = `${process.platform}:${ext || '<none>'}`;
  const cached = extCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now()) {
    return {
      path: resolved,
      ext,
      platform: process.platform,
      defaultApp: cached.defaultApp,
      apps: cached.apps,
    };
  }

  let apps: OpenWithApp[] = [];
  if (process.platform === 'win32') {
    try {
      apps = await discoverWin32Apps(ext);
    } catch {
      apps = [];
    }
  }

  const defaultApp = apps.find((a) => a.isDefault) ?? apps[0] ?? null;
  extCache.set(cacheKey, {
    expiresAt: Date.now() + CACHE_TTL_MS,
    apps,
    defaultApp,
  });

  return {
    path: resolved,
    ext,
    platform: process.platform,
    defaultApp,
    apps,
  };
}

export async function openExternalFile(
  rawPath: string,
  action: 'default' | 'app' | 'choose' | 'reveal',
  exePath?: string,
): Promise<{ ok: boolean; message: string }> {
  const resolved = resolveInputPath(rawPath);
  const stat = await fs.stat(resolved);
  if (!stat.isFile()) {
    throw new Error(`文件不存在: ${resolved}`);
  }

  if (action === 'default') {
    if (process.platform === 'win32') {
      spawn('explorer.exe', [resolved], { detached: true, stdio: 'ignore' }).unref();
    } else if (process.platform === 'darwin') {
      spawn('open', [resolved], { detached: true, stdio: 'ignore' }).unref();
    } else {
      spawn('xdg-open', [resolved], { detached: true, stdio: 'ignore' }).unref();
    }
    return { ok: true, message: `已调用系统默认程序打开 ${path.basename(resolved)}` };
  }

  if (action === 'app') {
    if (!exePath || !exePath.trim()) {
      throw new Error('缺少目标应用程序路径');
    }
    const targetExe = path.normalize(exePath.trim());
    if (process.platform === 'win32' && !targetExe.toLowerCase().endsWith('.exe')) {
      throw new Error('仅允许启动 Windows 可执行应用 (.exe)');
    }
    const exeKey = normalizeExeKey(targetExe);
    if (!allowedExePaths.has(exeKey)) {
      await listOpenWithApps(resolved);
      if (!allowedExePaths.has(exeKey)) {
        throw new Error(`未在系统注册列表中找到该应用: ${targetExe}`);
      }
    }
    const exeStat = await fs.stat(targetExe);
    if (!exeStat.isFile()) {
      throw new Error(`应用程序不存在: ${targetExe}`);
    }
    spawn(targetExe, [resolved], { detached: true, stdio: 'ignore', windowsHide: false }).unref();
    return { ok: true, message: `已使用 ${path.basename(targetExe, '.exe')} 打开 ${path.basename(resolved)}` };
  }

  if (action === 'choose') {
    if (process.platform === 'win32') {
      spawn('rundll32.exe', ['shell32.dll,OpenAs_RunDLL', resolved], {
        detached: true,
        stdio: 'ignore',
      }).unref();
      return { ok: true, message: '已打开系统「选择打开方式」对话框' };
    }
    return { ok: false, message: '当前操作系统不支持系统级「选择其他应用」对话框' };
  }

  // reveal
  if (process.platform === 'win32') {
    spawn('explorer.exe', [`/select,"${resolved}"`], {
      detached: true,
      stdio: 'ignore',
      windowsVerbatimArguments: true,
    }).unref();
  } else if (process.platform === 'darwin') {
    spawn('open', ['-R', resolved], { detached: true, stdio: 'ignore' }).unref();
  } else {
    spawn('xdg-open', [path.dirname(resolved)], { detached: true, stdio: 'ignore' }).unref();
  }
  return { ok: true, message: `已在文件管理器中定位 ${path.basename(resolved)}` };
}
