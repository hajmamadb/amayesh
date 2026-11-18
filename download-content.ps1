<#
.SYNOPSIS
  دانلود و ساماندهی کامل محتوای سایت‌های آماری
  (amar.org.ir، iranstatis.com، imidro.gov.ir، behinyab.ir، mimt.gov.ir، maj.ir)

.DESCRIPTION
  این اسکریپت:
  1) سایت مرکز آمار ایران (amar.org.ir - پلتفرم DNN):
     - تمام صفحات با Crawl (BFS) از صفحه اصلی
     - تمام اخبار از طریق RSS فید (چند هزار خبر)
     - تمام فایل‌های پیوست: PDF, Word, Excel, ZIP, تصاویر, صوت/ویدیو (پادکست), CSS/JS
  2) سایت آمار ایران (iranstatis.com - پلتفرم WordPress):
     - تمام پست‌ها و صفحات از REST API (با پشتیبان Sitemap)
     - کل کتابخانه رسانه (تصاویر, فونت‌ها, ویدیوها, فایل‌های داده)
     - داده‌های «آزمایشگاه داده» از API اختصاصی (سری‌های زمانی آمار)
     - متادیتا کامل (دسته‌ها, برچسب‌ها, فهرست منابع)
  3) سایر سایت‌های درخواستی با Crawl عمومی (BFS) از صفحه اصلی + صفحات مشخص‌شده:
     - imidro.gov.ir: گزارش‌های عملکرد، گزارش‌های آماری، آمار تولیدات سالانه، طرح‌ها و پروژه‌ها
     - behinyab.ir: درخت طبقه‌بندی ISIC (انواع فعالیت اقتصادی) + سایر بخش‌ها
     - mimt.gov.ir: آمارهای رسمی (معدن/صنعت/تجارت)، گزارش سالانه، گزارشات،
       موزه‌های علمی، محصولات، کتابچه راهنمای صنعت، مقررات/اکتشافات، سامانه‌ها
     - maj.ir: داشبورد اطلاعاتی و سامانه‌های کاربردی + سایر بخش‌ها
  4) گزارش و مانیفست (manifest.csv) برای هر سایت + خلاصه نهایی

  اسکریپت قابل اجرا مجدد است (Resume): فایل‌های قبلاً دانلودشده تکرار نمی‌شوند.

.PARAMETER OutputRoot
  پوشه مقصد (پیش‌فرض: D:\amayesh\محتوای جدید ساماندهی)

.PARAMETER Site
  سایت هدف: all | both | amar | iranstatis | imidro | behinyab | mimt | maj (پیش‌فرض: all)
  - all: همهٔ سایت‌ها
  - both: فقط amar.org.ir و iranstatis.com (رفتار قبلی)

.PARAMETER Workers
  تعداد درخواست‌های هم‌زمان (پیش‌فرض: 4)

.PARAMETER RatePerSec
  حداکثر نرخ درخواست کل (پیش‌فرض: 3 — برای جلوگیری از محدودسازی توسط سرور)

.PARAMETER MaxPages
  سقف تعداد صفحات قابل Crawl برای هر سایت (پیش‌فرض: 10000)

.PARAMETER ProToken
  توکن اشتراک حرفه‌ای iranstatis.com (اختیاری؛ برای ۸۵ دادهٔ غیرعمومی)

.PARAMETER NoResume
  شروع از صفر ( نادیده‌گرفتن مانیفست قبلی )

.EXAMPLE
  pwsh -File .\download-content.ps1                      (همهٔ سایت‌ها)
  pwsh -File .\download-content.ps1 -Site amar
  pwsh -File .\download-content.ps1 -Site imidro -Workers 6
  pwsh -File .\download-content.ps1 -Site iranstatis -ProToken "xxxx"
#>
[CmdletBinding()]
param(
    [string]$OutputRoot = 'D:\amayesh\محتوای جدید ساماندهی',
    [ValidateSet('all', 'both', 'amar', 'iranstatis', 'imidro', 'behinyab', 'mimt', 'maj')][string]$Site = 'all',
    [int]$Workers = 4,
    [double]$RatePerSec = 3,
    [int]$MaxPages = 10000,
    [string]$ProToken = '',
    [switch]$NoResume
)

# --- Configuration of extra sites (generic crawler: homepage/seed BFS crawl) ---
$script:ExtraSites = [ordered]@{
    'imidro'   = @{ Host = 'imidro.gov.ir';   Seeds = @(
        'https://imidro.gov.ir/',
        'https://imidro.gov.ir/fa/page/102472-%DA%AF%D8%B2%D8%A7%D8%B1%D8%B4-%D8%B9%D9%85%D9%84%DA%A9%D8%B1%D8%AF-%D8%B3%D8%A7%D8%B2%D9%85%D8%A7%D9%86.html',
        'https://imidro.gov.ir/fa/page/102535-%DA%AF%D8%B2%D8%A7%D8%B1%D8%B4-%D9%87%D8%A7%DB%8C-%D8%A2%D9%85%D8%A7%D8%B1%DB%8C.html',
        'https://imidro.gov.ir/fa/page/102554-%D8%A2%D9%85%D8%A7%D8%B1-%D8%AA%D9%88%D9%84%DB%8C%D8%AF%D8%A7%D8%AA-%D8%B3%D8%A7%D9%84%D8%A7%D9%86%D9%87-%D8%A8%D9%87-%D8%AA%D9%81%DA%A9%DB%8C%DA%A9-%DA%AF%D8%B1%D9%88%D9%87-%D9%85%D8%AD%D8%B5%D9%88%D9%84.html',
        'https://imidro.gov.ir/fa/page/52597-%D8%B7%D8%B1%D8%AD-%D9%87%D8%A7-%D9%BE%D8%B1%D9%88%DA%98%D9%87-%D9%87%D8%A7.html'
    ) }
    'behinyab' = @{ Host = 'behinyab.ir';     Seeds = @(
        'https://behinyab.ir/',
        'https://behinyab.ir/common/isic/tree'
    ) }
    'mimt'     = @{ Host = 'www.mimt.gov.ir'; Seeds = @(
        'https://www.mimt.gov.ir/',
        'https://www.mimt.gov.ir/%D8%A2%D9%85%D8%A7%D8%B1%D9%87%D8%A7%DB%8C-%D8%B1%D8%B3%D9%85%DB%8C/%D9%85%D8%B9%D8AF%D9%86',
        'https://www.mimt.gov.ir/%D8%A2%D9%85%D8%A7%D8%B1%D9%87%D8%A7%DB%8C-%D8%B1%D8%B3%D9%85%DB%8C/%D8%B5%D9%86%D8%B9%D8%AA',
        'https://www.mimt.gov.ir/%D8%A2%D9%85%D8%A7%D8%B1%D9%87%D8%A7%DB%8C-%D8%B1%D8%B3%D9%85%DB%8C/%D8%AA%D8%AC%D8%A7%D8%B1%D8%AA',
        'https://www.mimt.gov.ir/%DA%AF%D8%B2%D8%A7%D8%B1%D8%B4-%D8%B3%D8%A7%D9%84%D8%A7%D9%86%D9%87',
        'https://www.mimt.gov.ir/%DA%AF%D8%B2%D8%A7%D8%B1%D8%B4%D8%A7%D8%AA',
        'https://www.mimt.gov.ir/%D9%85%D9%88%D8%B6%D9%88%D8%B9%D8%A7%D8%AA-%D8%AF%D8%A7%D8%B1%D8%A7%DB%8C-%D8%A7%D9%88%D9%84%D9%88%DB%8C%D8%AA',
        'https://www.mimt.gov.ir/%D9%85%D8%AD%D8%B5%D9%88%D8%B1%D9%84%D8%A7%D8%AA-%D8%AF%D8%A7%D8%B1%D8%A7%DB%8C-%D8%A7%D9%88%D9%84%D9%88%DB%8C%D8%AA',
        'https://www.mimt.gov.ir/%DA%98%D8%AA%D8%A7%D8%A8%DA%86%D9%87-%D9%87%D8%A7%DB%8C-%D8%B1%D8%A7%D9%87%D9%86%D9%85%D8%A7%DB%8C-%D8%B5%D9%86%D8%B9%D8%AA',
        'https://www.mimt.gov.ir/regulations/%D8%A7%DA%A9%D8%AA%D8%B4%D8%A7%D9%81%D8%A7%D8%AA-%D9%88-%D8%A8%D9%87%D8%B1%D9%87-%D8%A8%D8%B1%D8%AF%D8%A7%D8%B1%DB%8C-%D9%85%D8%B9%D8%AF%D9%86',
        'https://www.mimt.gov.ir/%D8%B3%D8%A7%D9%85%D8%A7%D9%86%D9%87%E2%80%8C%D9%87%D8%A7'
    ) }
    'maj'      = @{ Host = 'maj.ir';          Seeds = @(
        'https://maj.ir/',
        'https://maj.ir/%D8%AF%D8%A7%D8%B4%D8%A8%D9%88%D8%B1%D8%AF-%D8%A7%D8%B7%D9%84%D8%A7%D8%B9%D8%A7%D8%AA%DB%8C-%D9%88-%D8%B3%D8%A7%D9%85%D8%A7%D9%86%D9%87-%D9%87%D8%A7%DB%8C-%DA%A9%D8%A7%D8%B1%D8%A8%D8%B1%D8%AF%DB%8C'
    ) }
}

#region --- Setup ---
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
try { $OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
try { Add-Type -AssemblyName System.Net.Http } catch {}

$LogDir = Join-Path $OutputRoot 'گزارش'
$LogFile = Join-Path $LogDir 'log.txt'

function Write-Log([string]$Msg, [string]$Level = 'INFO') {
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Msg
    try { Add-Content -Path $LogFile -Value $line -Encoding utf8 -ErrorAction SilentlyContinue } catch {}
    $color = switch ($Level) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } 'DONE' { 'Green' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
}

foreach ($d in @($OutputRoot, $LogDir,
    (Join-Path $OutputRoot 'amar.org.ir\گزارش'),
    (Join-Path $OutputRoot 'iranstatis.com\گزارش'))) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
foreach ($k in $script:ExtraSites.Keys) {
    $d = Join-Path $OutputRoot ("{0}\گزارش" -f $script:ExtraSites[$k].Host)
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

# --- HTTP client ---
$Handler = [System.Net.Http.HttpClientHandler]::new()
if ($Handler.PSObject.Properties.Name -contains 'AllowAutomaticDecompression') {
    try { $Handler.AllowAutomaticDecompression = [System.Net.DecompressionMethods]::All } catch {}
}
try { $Handler.MaxConnectionsPerServer = 12 } catch {}
$script:Hc = [System.Net.Http.HttpClient]::new($Handler)
$script:Hc.Timeout = [TimeSpan]::FromSeconds(90)
try {
    $script:Hc.DefaultVersion = [System.Net.HttpVersion]::Version11
    $script:Hc.DefaultVersionPolicy = [System.Net.HttpVersionPolicy]::RequestVersionExact
} catch {}
$script:Hc.DefaultRequestHeaders.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36')
$script:Hc.DefaultRequestHeaders.Add('Accept', 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8')
$script:Hc.DefaultRequestHeaders.Add('Accept-Language', 'fa-IR,fa;q=0.9,en;q=0.7')

# --- Global rate limiter ---
$script:RateLock = New-Object object
$script:RateNextMs = [long]0
function Get-NextSlotMs([long]$MinIntervalMs) {
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    [System.Threading.Monitor]::Enter($script:RateLock)
    try {
        $slot = [Math]::Max($now, $script:RateNextMs)
        $script:RateNextMs = $slot + $MinIntervalMs
        return $slot
    } finally { [System.Threading.Monitor]::Exit($script:RateLock) }
}
$script:MinIntervalMs = [long][Math]::Max(20, [int](1000 / [Math]::Max(0.5, $RatePerSec)))

# --- Anti-block cooldown (server WAF intermittently returns 404 after bursts) ---
$script:FailStreak = [int]0
$script:RecentFails = [int]0
function Note-HttpOutcome([bool]$Success, [bool]$Blocking = $true) {
    if ($Success) { $script:FailStreak = 0; $script:RecentFails = 0; return }
    if (-not $Blocking) { return }
    $script:FailStreak += 1
    $script:RecentFails += 1
    if ($script:FailStreak -ge 10) {
        Write-Log ("سرور همچنان محدودسازی می‌کند — توقف ۱۰ دقیقه‌ای (خطای متوالی: {0})" -f $script:FailStreak) 'WARN'
        Start-Sleep -Seconds 600
        $script:FailStreak = 0
    }
    elseif ($script:FailStreak -ge 5) {
        Write-Log ("سرور در حال محدودسازی دسترسی است — توقف ۳ دقیقه‌ای (خطای متوالی: {0})" -f $script:FailStreak) 'WARN'
        Start-Sleep -Seconds 180
        $script:FailStreak = 0
    }
}
function In-BlockWindow { $script:RecentFails -ge 3 }

# --- Counters ---
$script:Stats = [pscustomobject]@{ Done = [long]0; Errors = [long]0; Bytes = [long]0; StatsLock = (New-Object object) }
function Bump-Stat([string]$Field, [long]$Delta = 1) {
    [System.Threading.Monitor]::Enter($script:Stats.StatsLock)
    try { $script:Stats.$Field += $Delta } finally { [System.Threading.Monitor]::Exit($script:Stats.StatsLock) }
}

# --- Synchronous fetch with retries (for feeds / REST API on main thread) ---
function Fetch-Url([string]$Url, [int]$Tries = 4) {
    for ($t = 1; $t -le $Tries; $t++) {
        try {
            $slot = Get-NextSlotMs $script:MinIntervalMs
            $wait = $slot - [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
            if ($wait -gt 0) { [Threading.Thread]::Sleep([int]$wait) }
            $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $Url)
            $resp = $script:Hc.SendAsync($req, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
            $code = [int]$resp.StatusCode
            $ct = ''
            if ($resp.Content.Headers.ContentType) { $ct = $resp.Content.Headers.ContentType.ToString() }
            if ($code -ge 400) {
                $resp.Dispose()
                $blocking = ($code -eq 403 -or $code -eq 429 -or $code -ge 500 -or ($code -eq 404 -and (In-BlockWindow)))
                Note-HttpOutcome $false $blocking
                # For API/feed calls, treat 404/403/429/5xx as transient (server-side rate limiting is observed)
                if ($code -eq 403 -or $code -eq 404 -or $code -eq 429 -or $code -ge 500) {
                    if ($t -lt $Tries) { Start-Sleep -Seconds ([Math]::Min(30, 2 * $t * $t)); continue }
                }
                return [pscustomobject]@{ Bytes = $null; Status = $code; ContentType = $ct; Error = "HTTP $code" }
            }
            Note-HttpOutcome $true
            $ms = [IO.MemoryStream]::new()
            $stream = $resp.Content.ReadAsStream()
            $stream.CopyTo($ms)
            $stream.Dispose()
            $resp.Dispose()
            $ms.Position = 0
            $bytes = New-Object byte[] $ms.Length
            [void]$ms.Read($bytes, 0, [int]$ms.Length)
            $ms.Dispose()
            return [pscustomobject]@{ Bytes = $bytes; Status = $code; ContentType = $ct; Error = '' }
        } catch {
            $emsg = $_.Exception.Message
            $isNetErr = ($emsg -match '(?i)host|dns|name resolution|connect|reset|socket|timeout|timed|stream')
            Note-HttpOutcome $false $isNetErr
            if ($t -lt $Tries) {
                if ($isNetErr) {
                    Write-Log "خطای شبکه/DNS برای $Url — انتظار ۵ دقیقه‌ای قبل از تلاش بعدی" 'WARN'
                    Start-Sleep -Seconds 300
                } else { Start-Sleep -Seconds ([Math]::Min(30, 2 * $t * $t)) }
                continue
            }
            return [pscustomobject]@{ Bytes = $null; Status = 0; ContentType = ''; Error = $emsg }
        }
    }
}
#endregion

#region --- Manifest ---
$script:Manifest = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
$script:ManifestLock = New-Object object
$script:PathOwner = [System.Collections.Concurrent.ConcurrentDictionary[string, string]]::new()

function New-Record([string]$Url, [string]$Kind, [string]$Site, [int]$Status, [long]$Size, [string]$LocalPath, [string]$Err) {
    [pscustomobject]@{
        Site = $Site; Type = $Kind; URL = $Url; HTTPStatus = $Status; SizeBytes = $Size
        LocalPath = $LocalPath; Error = $Err; DownloadedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
}

function Save-Record([string]$Key, [object]$Rec) {
    $null = $script:Manifest.AddOrUpdate($Key, $Rec, { param($k, $o) $Rec })
}

function Export-ManifestCsv([string]$CsvPath, [string]$Site) {
    $rows = @()
    foreach ($k in $script:Manifest.Keys) {
        $rec = $script:Manifest[$k]
        if ($rec.Site -eq $Site) { $rows += $rec }
    }
    $lines = @('Site,Type,URL,HTTPStatus,SizeBytes,LocalPath,Error,DownloadedAt')
    foreach ($r in $rows) {
        $lines += ('{0},{1},"{2}",{3},{4},"{5}","{6}","{7}"' -f $r.Site, $r.Type, ($r.URL -replace '"', '""'), $r.HTTPStatus, $r.SizeBytes, ($r.LocalPath -replace '"', '""'), ($r.Error -replace '"', '""'), $r.DownloadedAt)
    }
    $tmp = "$CsvPath.tmp"
    [System.IO.File]::WriteAllLines($tmp, $lines, [System.Text.UTF8Encoding]::new($true))
    try {
        if (Test-Path $CsvPath) { Remove-Item $CsvPath -Force }
        Move-Item $tmp $CsvPath -Force
    } catch {
        if (Test-Path $tmp) { Remove-Item $tmp -Force }
    }
}

function Load-ExistingManifest([string]$Site) {
    $csv = Join-Path (Join-Path $OutputRoot "$Site\گزارش") 'manifest.csv'
    if ((Test-Path $csv) -and (-not $NoResume)) {
        try {
            $reader = [System.IO.File]::CreateTextReader($csv, [System.Text.Encoding]::UTF8)
            $null = $reader.ReadLine()
            $cnt = 0
            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                if (-not $line) { continue }
                $m = [regex]::Match($line, '^([^,]*),([^,]*),\s*"((?:[^"]|"")*)",\s*(\d+)')
                if ($m.Success) {
                    $url = $m.Groups[3].Value -replace '""', '"'
                    $st = [int]$m.Groups[4].Value
                    if ($st -eq 200) {
                        $script:Manifest[$url] = New-Record $url 'existing' $Site 200 0 '(cached)' ''
                        $cnt++
                    }
                }
            }
            $reader.Close()
            Write-Log "Resume: $cnt آدرس از مانیفست قبلی بارگذاری شد ($Site)"
        } catch { Write-Log "خطا در خواندن مانیفست: $($_.Exception.Message)" 'WARN' }
    }
}
#endregion

#region --- Path helpers ---
$InvalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
$InvalidRegex = [regex]::Escape($InvalidChars)

function Get-SafeSegment([string]$Seg) {
    $seg = [System.Uri]::UnescapeDataString($Seg)
    $seg = $seg -replace "[$InvalidRegex]", '_'
    $seg = $seg -replace '[\x00-\x1f]', ''
    $seg = $seg.Trim().TrimEnd('.')
    if ($seg.Length -eq 0) { return '_' }
    if ($seg -match '^(con|prn|aux|nul|com[1-9]|lpt[1-9])$') { $seg = "_$seg" }
    if ($seg.Length -gt 120) {
        $hash = (Get-FileHash -InputStream ([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($seg))) -Algorithm MD5).Hash.ToLower()
        $seg = $seg.Substring(0, 100) + '_' + $hash.Substring(0, 8)
    }
    return $seg
}

function Get-SafeFileName([string]$Name) {
    $n = $Name -replace "[$InvalidRegex]", '_'
    $n = $n -replace '[\x00-\x1f]', ''
    $n = $n.Trim().TrimEnd('.')
    if ($n.Length -eq 0) { $n = 'unnamed' }
    if ($n.Length -gt 150) {
        $hash = (Get-FileHash -InputStream ([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($Name))) -Algorithm MD5).Hash.ToLower()
        $ext = [IO.Path]::GetExtension($n)
        $base = $n.Substring(0, [Math]::Max(0, $n.Length - $ext.Length))
        $n = $base.Substring(0, [Math]::Min(140, $base.Length)) + '_' + $hash.Substring(0, 8) + $ext
    }
    return $n
}

$KeepParams = @('catid', 'categoryid', 'tabid', 'id', 'page', 'ctl', 'articletype', 'tag')
function Get-UrlQuerySuffix([string]$Query) {
    if (-not $Query) { return '' }
    $q = $Query.TrimStart('?')
    $parts = @()
    foreach ($kv in $q.Split('&')) {
        if (-not $kv) { continue }
        $i = $kv.IndexOf('=')
        if ($i -lt 0) { $k = $kv; $v = '' } else { $k = $kv.Substring(0, $i); $v = $kv.Substring($i + 1) }
        if ($KeepParams -notcontains $k.ToLower()) { continue }
        if ($v.Length -gt 40) { $v = $v.Substring(0, 40) }
        $parts += ($k.ToLower() -replace "[$InvalidRegex]", '_') + '-' + ($v -replace "[$InvalidRegex]", '_')
        if ($parts.Count -ge 3) { break }
    }
    if ($parts.Count -eq 0) { return '' }
    $s = ($parts -join '_')
    if ($s.Length -gt 80) { $s = $s.Substring(0, 80) }
    return '_' + $s
}

function Get-UrlPathSegments([string]$Url) {
    $u = [System.Uri]($Url)
    $path = $u.AbsolutePath.TrimEnd('/')
    if ($path.Length -eq 0) { return @() }
    $segs = $path.Split('/') | Where-Object { $_.Length -gt 0 }
    return @($segs | ForEach-Object { Get-SafeSegment $_ })
}
#endregion

#region --- Classification ---
$DocExts   = @('pdf', 'doc', 'docx', 'xls', 'xlsx', 'xlsm', 'ppt', 'pptx', 'zip', 'rar', '7z', 'csv', 'dbf', 'shp', 'shx', 'prj', 'geojson', 'json', 'odt', 'ods', 'odp', 'epub', 'txt', 'rtf')
$ImgExts   = @('jpg', 'jpeg', 'png', 'gif', 'svg', 'webp', 'ico', 'bmp', 'avif', 'tiff', 'tif', 'heic')
$MediaExts = @('mp4', 'mp3', 'wav', 'webm', 'm4a', 'aac', 'ogg', 'oga', 'mov', 'flv', 'wmv', 'mkv')
$AssetExts = @('css', 'js', 'woff', 'woff2', 'ttf', 'otf', 'eot', 'map')
$FeedExts  = @('xml', 'rss', 'atom')

function Get-Ext([string]$Url) {
    try { $p = [System.Uri]$Url | ForEach-Object { $_.AbsolutePath } }
    catch { $p = ($Url -split '\?')[0] }
    if (-not $p) { return '' }
    if ($p.EndsWith('/')) { $p = $p.TrimEnd('/') }
    $e = [IO.Path]::GetExtension($p)
    return $e.TrimStart('.').ToLower()
}

function Get-UrlKind([string]$Url) {
    $e = Get-Ext $Url
    $u = $Url.ToLower()
    if ($e -in $DocExts)   { return 'doc' }
    if ($e -in $ImgExts)   { return 'image' }
    if ($e -in $MediaExts) { return 'media' }
    if ($e -in $FeedExts)  { return 'feed' }
    if ($e -in $AssetExts) { return 'asset' }
    if ($e -in @('ashx', 'aspx', 'php', 'asp')) {
        if ($u -match 'rss|feed') { return 'feed' }
        return 'handler'
    }
    if ($u -match '\.(html|htm)$') { return 'page' }
    if ($e -eq '') { return 'page' }
    return 'other'
}

function Get-KindFolder([string]$Site, [string]$Kind) {
    $base = Join-Path $OutputRoot $Site
    switch ($Kind) {
        'page'    { Join-Path $base 'صفحات' }
        'doc'     { Join-Path $base 'اسناد' }
        'image'   { Join-Path $base 'تصاویر' }
        'media'   { Join-Path $base 'چندرسانه' }
        'asset'   { Join-Path $base 'کدواستایل' }
        'feed'    { Join-Path $base 'فیدها' }
        'data'    { Join-Path $base 'داده‌ها' }
        'apidata' { Join-Path $base 'داده-آزمایشگاه' }
        default   { Join-Path $base 'سایر' }
    }
}

function Resolve-LocalPath([string]$Site, [string]$Kind, [string]$Url, [string]$ForceExt = '') {
    $root = Get-KindFolder $Site $Kind
    $u = [System.Uri]($Url)
    $q = $u.Query.TrimStart('?')
    if ($u.AbsolutePath -match '(?i)imagehandler|download\.php|get\.php|attachment\.php') {
        $leaf = (Split-Path $u.AbsolutePath -Leaf).Replace('.ashx', '').Replace('.php', '')
        $hash = (Get-FileHash -InputStream ([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($Url))) -Algorithm MD5).Hash.ToLower()
        $name = Get-SafeFileName ("{0}_{1}{2}" -f $leaf, $hash.Substring(0, 10), $ForceExt)
        $full = Join-Path $root $name
    }
    else {
        $segs = Get-UrlPathSegments $Url
        if ($segs.Count -eq 0) {
            $full = Join-Path $root 'index.html'
        }
        else {
            $last = $segs[-1]
            $e = Get-Ext $Url
            if ($ForceExt) { $ext = $ForceExt } elseif ($e) { $ext = '.' + $e } else { $ext = '.html' }
            $suffix = Get-UrlQuerySuffix $q
            $last = Get-SafeFileName ($last + $suffix + $ext)
            $dir = $root
            if ($segs.Count -gt 1) { $dir = Join-Path $root (($segs | Select-Object -SkipLast 1) -join '\') }
            $full = Join-Path $dir $last
        }
    }
    $dirOf = Split-Path $full -Parent
    if (-not (Test-Path $dirOf)) { New-Item -ItemType Directory -Path $dirOf -Force | Out-Null }
    return $full
}

function Resolve-UniquePath([string]$Path, [string]$Url) {
    $owner = $null
    if ($script:PathOwner.TryGetValue($Path, [ref]$owner)) {
        if ($owner -eq $Url) { return $Path }
    }
    if (-not (Test-Path $Path)) { return $Path }
    $hash = (Get-FileHash -InputStream ([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($Url))) -Algorithm MD5).Hash.ToLower().Substring(0, 8)
    $dir = Split-Path $Path -Parent
    $n = [IO.Path]::GetFileNameWithoutExtension($Path)
    $e = [IO.Path]::GetExtension($Path)
    return (Join-Path $dir ($n + '_' + $hash + $e))
}
#endregion

#region --- Crawl engine (async HTTP, single-threaded control loop) ---
$script:WorkQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
$script:PagesSeen = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
$script:AssetsSeen = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
$script:PagesCount = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
$script:StopCrawl = $false

function New-Item2([string]$Url, [string]$Kind, [string]$Site, [int]$Depth = 0, [int]$Retry = 0) {
    [pscustomobject]@{ Url = $Url; Kind = $Kind; Site = $Site; Depth = $Depth; Retry = $Retry }
}

function Enqueue-Asset([string]$Url, [string]$Kind, [string]$Site) {
    if ($script:StopCrawl) { return }
    $existingRec = $null
    if ($script:Manifest.TryGetValue($Url, [ref]$existingRec)) { return }  # resume: already downloaded
    if ($script:AssetsSeen.TryAdd($Url, 1)) {
        $script:WorkQueue.Enqueue((New-Item2 $Url $Kind $Site))
    }
}

function Enqueue-Page([string]$Url, [string]$Site, [int]$Depth) {
    if ($script:StopCrawl) { return }
    if ($Depth -gt 14) { return }
    if ($script:PagesCount.Count -ge $MaxPages) { return }
    $existingRec = $null
    if ($script:Manifest.TryGetValue($Url, [ref]$existingRec)) { return }  # resume: already downloaded
    if ($script:PagesSeen.TryAdd($Url, 1)) {
        $script:PagesCount[$Url] = 1
        $script:WorkQueue.Enqueue((New-Item2 $Url 'page' $Site $Depth))
    }
}

function Get-HtmlUrls([string]$Html) {
    $found = New-Object System.Collections.Generic.List[string]
    $tags = [regex]::Matches($Html, '(?is)<(a|area|img|source|video|audio|track|embed|iframe|link|script)\b[^>]*>')
    foreach ($t in $tags) {
        $m = [regex]::Match($t.Value, '(?i)\b(?:href|src|poster|data-src)\s*=\s*"([^"]+)"')
        if (-not $m.Success) { $m = [regex]::Match($t.Value, "(?i)\b(?:href|src|poster|data-src)\s*=\s*'([^']+)") }
        if ($m.Success) {
            $v = $m.Groups[1].Value.Trim()
            if ($v.Length -gt 0 -and $v -notmatch '^(javascript:|#|data:|about:)') { $found.Add($v) }
        }
    }
    return $found
}

function Normalize-Url([string]$Base, [string]$Rel) {
    if ($Rel -match '^(javascript:|#|data:|about:|mailto:|tel:)') { return $null }
    try {
        $abs = [System.Uri]::new([System.Uri]$Base, $Rel)
        if ($abs.Scheme -notin @('http', 'https')) { return $null }
        $auth = $abs.Authority.ToLower()
        $u = 'https://' + $auth + $abs.AbsolutePath
        if ($abs.Query) { $u += $abs.Query }
        return $u
    } catch { return $null }
}

function Clean-PageUrl([string]$Url) {
    try {
        $u = [System.Uri]$Url
        $q = $u.Query.TrimStart('?')
        if (-not $q) { return $Url }
        $keep = @()
        foreach ($kv in $q.Split('&')) {
            if (-not $kv) { continue }
            $k = ($kv.Split('='))[0].ToLower()
            if ($k -in @('cdv', 'returnurl', 'utm_source', 'utm_medium', 'utm_campaign', 'utm_term', 'utm_content', 'fbclid', 'gclid', 'msclkid', 'nonce', 'wp_customize')) { continue }
            $keep += $kv
        }
        $newq = ($keep | Sort-Object -Unique) -join '&'
        if ($newq.Length -eq 0) { return ($Url -replace '\?.*$', '') }
        return ($Url -replace '\?.*$', '?' + $newq)
    } catch { return $Url }
}

function Detect-KindFromBytes([byte[]]$B) {
    if ($B.Length -lt 12) { return 'other' }
    if ($B[0] -eq 0x25 -and $B[1] -eq 0x50) { return 'doc' }
    if ($B[0] -eq 0xD0 -and $B[1] -eq 0xCF) { return 'doc' }
    if ($B[0] -eq 0x50 -and $B[1] -eq 0x4B) { return 'doc' }
    if ($B[0] -eq 0xFF -and $B[1] -eq 0xD8) { return 'image' }
    if ($B[0] -eq 0x89 -and $B[1] -eq 0x50) { return 'image' }
    if ($B[0] -eq 0x47 -and $B[1] -eq 0x49) { return 'image' }
    if ($B[0] -eq 0x42 -and $B[1] -eq 0x4D) { return 'image' }
    if ($B[8] -eq 0x57 -and $B[9] -eq 0x45 -and $B[10] -eq 0x42) { return 'image' }
    if ($B[0] -eq 0x49 -and $B[1] -eq 0x44) { return 'media' }
    if ($B[4] -eq 0x66 -and $B[5] -eq 0x74) { return 'media' }
    $head = [Text.Encoding]::ASCII.GetString($B, 0, [Math]::Min(200, $B.Length))
    if ($head -match '(?i)<\?xml|<rss|<feed') { return 'feed' }
    if ($head -match '(?is)<(html|head|body|!doctype)') { return 'page' }
    return 'other'
}

function Guess-ImageExt([byte[]]$B) {
    if ($B.Length -lt 12) { return '' }
    if ($B[0] -eq 0xFF -and $B[1] -eq 0xD8) { return '.jpg' }
    if ($B[0] -eq 0x89 -and $B[1] -eq 0x50) { return '.png' }
    if ($B[0] -eq 0x47 -and $B[1] -eq 0x49) { return '.gif' }
    if ($B[0] -eq 0x42 -and $B[1] -eq 0x4D) { return '.bmp' }
    if ($B[8] -eq 0x57 -and $B[9] -eq 0x45 -and $B[10] -eq 0x42) { return '.webp' }
    return ''
}

function Handle-PageResult([pscustomobject]$Item, [byte[]]$Bytes, [int]$Code) {
    $url = $Item.Url; $site = $Item.Site
    $up = $url.ToLower()
    if ($up -match '/(admin|app_code|app_data|app_browsers|bin|config|install|providers|httpmodules|components|documentation|controls|activity-feed)/') { return }
    if ($up -match '^https?://[^/]+/search(\?|$)') { return }
    if ($up -match '(?i)/default\.aspx') { return }
    if ($up -match '\?returnurl=') { return }

    $html = [Text.Encoding]::UTF8.GetString($Bytes)
    if ($html -match '(?is)<(html|head|body|!doctype)') {
        $path = Resolve-LocalPath $site 'page' $url
        $path = Resolve-UniquePath $path $url
        [IO.File]::WriteAllBytes($path, $Bytes)
        $script:PathOwner[$path] = $url
        Save-Record $url (New-Record $url 'page' $site $Code $Bytes.Length $path '')
        Bump-Stat 'Done'; Bump-Stat 'Bytes' $Bytes.Length

        $host_ = ([System.Uri]$url).Authority
        foreach ($rel in (Get-HtmlUrls $html)) {
            $abs = Normalize-Url $url $rel
            if (-not $abs) { continue }
            if (([System.Uri]$abs).Authority -ne $host_) { continue }
            $al = $abs.ToLower()
            if ($al -match '^https?://[^/]+/(resources|app_|admin|bin|config|install|providers|httpmodules)/') { continue }
            $kind = Get-UrlKind $abs
            if ($kind -eq 'page') {
                Enqueue-Page (Clean-PageUrl $abs) $site ($Item.Depth + 1)
            }
            else {
                Enqueue-Asset $abs $kind $site
            }
        }
    }
    else {
        $path = Resolve-LocalPath $site 'other' $url
        $path = Resolve-UniquePath $path $url
        [IO.File]::WriteAllBytes($path, $Bytes)
        $script:PathOwner[$path] = $url
        Save-Record $url (New-Record $url 'other' $site $Code $Bytes.Length $path '')
        Bump-Stat 'Done'; Bump-Stat 'Bytes' $Bytes.Length
    }
}

function Handle-AssetResult([pscustomobject]$Item, [byte[]]$Bytes, [int]$Code) {
    $url = $Item.Url; $kind = $Item.Kind; $site = $Item.Site
    $finalKind = $kind
    $forceExt = ''
    if ($kind -eq 'handler') {
        $finalKind = Detect-KindFromBytes $Bytes
        switch ($finalKind) {
            'image' { $forceExt = (Guess-ImageExt $Bytes); if (-not $forceExt) { $forceExt = '.jpg' } }
            'feed'  { $forceExt = '.xml' }
            'page'  { $finalKind = 'other' }
            default { $finalKind = 'other' }
        }
    }
    $path = Resolve-LocalPath $site $finalKind $url $forceExt
    $path = Resolve-UniquePath $path $url
    [IO.File]::WriteAllBytes($path, $Bytes)
    $script:PathOwner[$path] = $url
    Save-Record $url (New-Record $url $finalKind $site $Code $Bytes.Length $path '')
    Bump-Stat 'Done'; Bump-Stat 'Bytes' $Bytes.Length
}

function Invoke-CrawlEngine([string]$Site, [int]$MaxConcurrent) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $lastLog = [DateTime]::MinValue
    $lastManifestSave = Get-Date
    $inflight = New-Object System.Collections.Generic.List[object]

    while ($true) {
        # 1) start requests up to the concurrency limit
        while ($inflight.Count -lt $MaxConcurrent -and -not $script:StopCrawl) {
            $item = $null
            if (-not $script:WorkQueue.TryDequeue([ref]$item)) { break }
            $slot = Get-NextSlotMs $script:MinIntervalMs
            $wait = $slot - [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
            if ($wait -gt 0) { [Threading.Thread]::Sleep([int]$wait) }
            try {
                $req = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::Get, $item.Url)
                $task = $script:Hc.SendAsync($req, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead)
                $inflight.Add([pscustomobject]@{ Task = $task; Item = $item })
            } catch {
                Bump-Stat 'Errors'
                Save-Record $item.Url (New-Record $item.Url $item.Kind $Site 0 0 '' $_.Exception.Message)
            }
        }

        # 2) termination?
        if ($inflight.Count -eq 0) {
            if ($script:WorkQueue.Count -eq 0) { break }
            Start-Sleep -Milliseconds 200
            continue
        }

        # 3) wait for the next completion
        $taskArr = New-Object 'System.Threading.Tasks.Task[]' $inflight.Count
        for ($i = 0; $i -lt $inflight.Count; $i++) { $taskArr[$i] = $inflight[$i].Task }
        $idx = [System.Threading.Tasks.Task]::WaitAny($taskArr)
        if ($idx -lt 0) {
            # all cancelled/timed out
            $inflight.Clear()
            continue
        }
        $entry = $inflight[$idx]
        $inflight.RemoveAt($idx)

        try {
            $resp = $entry.Task.GetAwaiter().GetResult()
            $code = [int]$resp.StatusCode
            if ($code -ge 400) {
                $resp.Dispose()
                $isTransient = ($code -eq 403 -or $code -eq 429 -or $code -ge 500)
                $isBlocked404 = ($code -eq 404 -and (In-BlockWindow))
                Note-HttpOutcome $false ($isTransient -or $isBlocked404)
                if (($isTransient -or $isBlocked404) -and $entry.Item.Retry -lt 3) {
                    # transient / WAF block: requeue with retry
                    $script:WorkQueue.Enqueue((New-Item2 $entry.Item.Url $entry.Item.Kind $entry.Item.Site $entry.Item.Depth ($entry.Item.Retry + 1)))
                    $rv = $null
                    $null = $script:AssetsSeen.TryRemove($entry.Item.Url, [ref]$rv)
                    $null = $script:PagesSeen.TryRemove($entry.Item.Url, [ref]$rv)
                    $null = $script:PagesCount.TryRemove($entry.Item.Url, [ref]$rv)
                    continue
                }
                Save-Record $entry.Item.Url (New-Record $entry.Item.Url $entry.Item.Kind $Site $code 0 '' ("HTTP $code"))
                Bump-Stat 'Errors'
                continue
            }
            Note-HttpOutcome $true
            $ms = [IO.MemoryStream]::new()
            $st = $resp.Content.ReadAsStream()
            $st.CopyTo($ms)
            $st.Dispose()
            $resp.Dispose()
            $ms.Position = 0
            $bytes = New-Object byte[] $ms.Length
            [void]$ms.Read($bytes, 0, [int]$ms.Length)
            $ms.Dispose()
            if ($entry.Item.Kind -eq 'page') { Handle-PageResult $entry.Item $bytes $code }
            else { Handle-AssetResult $entry.Item $bytes $code }
        }
        catch {
            $emsg = $_.Exception.Message
            if ($entry.Item.Retry -lt 3 -and ($emsg -match '(?i)host|dns|name resolution|connect|reset|socket|timeout|timed|stream')) {
                # network-level failure (e.g. temporary DNS blackhole): requeue with retry
                Note-HttpOutcome $false $true
                $script:WorkQueue.Enqueue((New-Item2 $entry.Item.Url $entry.Item.Kind $entry.Item.Site $entry.Item.Depth ($entry.Item.Retry + 1)))
                $rv = $null
                $null = $script:AssetsSeen.TryRemove($entry.Item.Url, [ref]$rv)
                $null = $script:PagesSeen.TryRemove($entry.Item.Url, [ref]$rv)
                $null = $script:PagesCount.TryRemove($entry.Item.Url, [ref]$rv)
                continue
            }
            Bump-Stat 'Errors'
            Save-Record $entry.Item.Url (New-Record $entry.Item.Url $entry.Item.Kind $Site 0 0 '' $emsg)
        }

        # 4) progress log + periodic manifest save (so a stopped run still resumes correctly)
        $now = Get-Date
        if (($now - $lastLog).TotalSeconds -ge 20) {
            $lastLog = $now
            $done = $script:Stats.Done; $err = $script:Stats.Errors; $bz = $script:Stats.Bytes
            $rate = $done / [Math]::Max(1, $sw.Elapsed.TotalSeconds)
            Write-Log ("  [{0}] انجام‌شده={1} خطا={2} حجم={3:N0}MB صف={4} درحال‌دانلود={5} سرعت={6:N1}/s" -f $Site, $done, $err, ($bz / 1MB), $script:WorkQueue.Count, $inflight.Count, $rate)
            if (($now - $lastManifestSave).TotalMinutes -ge 5) {
                $lastManifestSave = $now
                try { Export-ManifestCsv (Join-Path (Join-Path $OutputRoot "$Site\گزارش") 'manifest.csv') $Site } catch {}
            }
        }
    }
    $sw.Stop()
    Write-Log ("[{0}] موتور crawl پایان یافت در {1:N0} ثانیه" -f $Site, $sw.Elapsed.TotalSeconds) 'DONE'
}
#endregion

#region --- Site: amar.org.ir ---
function Invoke-AmarCrawl {
    $site = 'amar.org.ir'
    Write-Log ("==== شروع crawl {0} ====" -f $site)
    $script:Stats = [pscustomobject]@{ Done = [long]0; Errors = [long]0; Bytes = [long]0; StatsLock = (New-Object object) }
    $script:WorkQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:PagesSeen = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $script:AssetsSeen = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $script:PagesCount = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $script:StopCrawl = $false
    Load-ExistingManifest $site

    # seed homepage
    $null = $script:PagesSeen.TryAdd('https://amar.org.ir/', 1)
    $script:PagesCount['https://amar.org.ir/'] = 1
    $script:WorkQueue.Enqueue((New-Item2 'https://amar.org.ir/' 'page' $site 0))

    # news RSS -> all news urls
    try {
        Write-Log 'دانلود فید خبرها (RSS) برای استخراج فهرست کامل اخبار...'
        $feedUrl = 'https://amar.org.ir/DesktopModules/DnnForge%20-%20NewsArticles/RSS.aspx?TabID=182&ModuleID=1576&MaxCount=5000'
        $r = Fetch-Url $feedUrl
        if ($r.Bytes) {
            $txt = [Text.Encoding]::UTF8.GetString($r.Bytes)
            $feedPath = Join-Path (Get-KindFolder $site 'feed') 'news-main-feed.xml'
            [IO.File]::WriteAllBytes($feedPath, $r.Bytes)
            Save-Record $feedUrl (New-Record $feedUrl 'feed' $site 200 $r.Bytes.Length $feedPath '')
            Bump-Stat 'Done'; Bump-Stat 'Bytes' $r.Bytes.Length
            $links = [regex]::Matches($txt, '(?is)<item>.*?<link>\s*(.*?)\s*</link>') | ForEach-Object { $_.Groups[1].Value.Trim() }
            $newsCount = 0
            foreach ($l in $links) {
                $a = Normalize-Url 'https://amar.org.ir/' $l
                if ($a -and $a -match '/news/id/') { Enqueue-Page $a $site 1; $newsCount++ }
            }
            Write-Log "فید خبر: $newsCount خبر در صف قرار گرفت" 'DONE'
        }
        else {
            Write-Log "فید خبر در دسترس نبود: $($r.Error)" 'WARN'
        }
    }
    catch {
        Write-Log "خطا در فید خبر: $($_.Exception.Message)" 'ERROR'
    }

    Invoke-CrawlEngine $site $Workers

    Export-ManifestCsv (Join-Path (Join-Path $OutputRoot "$site\گزارش") 'manifest.csv') $site
    Write-SiteSummary $site
}
#endregion

#region --- Site: iranstatis.com ---
function Invoke-IranstatisCrawl {
    $site = 'iranstatis.com'
    Write-Log ("==== شروع crawl {0} ====" -f $site)
    $script:Stats = [pscustomobject]@{ Done = [long]0; Errors = [long]0; Bytes = [long]0; StatsLock = (New-Object object) }
    $script:WorkQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:PagesSeen = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $script:AssetsSeen = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $script:PagesCount = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $script:StopCrawl = $false
    Load-ExistingManifest $site

    $metaDir = Join-Path (Get-KindFolder $site 'apidata') 'meta'
    if (-not (Test-Path $metaDir)) { New-Item -ItemType Directory -Path $metaDir -Force | Out-Null }

    function Read-Api([string]$Url) {
        $r = Fetch-Url $Url
        if (-not $r.Bytes) { throw $r.Error }
        return $r.Bytes
    }

    # ---- 1) REST API: pages, posts, media, categories, tags
    Write-Log 'استخراج فهرست کامل از REST API وردپرس...'
    $rest = @{ pages = @(); posts = @(); media = @(); categories = @(); tags = @() }
    foreach ($ep in @('pages', 'posts', 'categories', 'tags')) {
        $p = 1
        while ($true) {
            $url = "https://iranstatis.com/wp-json/wp/v2/$ep?per_page=100&page=$p"
            try { $b = Read-Api $url } catch { Write-Log "خطا REST $ep page $p : $($_.Exception.Message)" 'WARN'; break }
            $j = $b | ConvertFrom-Json
            if ($null -eq $j -or @($j).Count -eq 0) { break }
            foreach ($it in $j) { $rest[$ep] += $it }
            if (@($j).Count -lt 100) { break }
            $p++
            if ($p -gt 40) { break }
        }
        Write-Log ("  {0}: {1} مورد" -f $ep, $rest[$ep].Count)
    }
    $p = 1
    while ($true) {
        $url = "https://iranstatis.com/wp-json/wp/v2/media?per_page=100&page=$p"
        try { $b = Read-Api $url } catch { Write-Log "خطا REST media page $p : $($_.Exception.Message)" 'WARN'; break }
        $j = $b | ConvertFrom-Json
        if ($null -eq $j -or @($j).Count -eq 0) { break }
        foreach ($it in $j) { $rest['media'] += $it }
        if (@($j).Count -lt 100) { break }
        $p++
        if ($p -gt 30) { break }
    }
    Write-Log ("  media: {0} مورد" -f $rest['media'].Count)

    foreach ($k in @('pages', 'posts', 'categories', 'tags', 'media')) {
        $p2 = Join-Path $metaDir "$k.json"
        $jsonOut = $rest[$k] | ConvertTo-Json -Depth 6
        if (-not $jsonOut) { $jsonOut = '[]' }
        [IO.File]::WriteAllBytes($p2, [Text.Encoding]::UTF8.GetBytes($jsonOut))
    }

    foreach ($pg in $rest['pages']) {
        $u = Normalize-Url 'https://iranstatis.com/' $pg.link
        if ($u) { Enqueue-Page $u $site 0 }
    }
    foreach ($po in $rest['posts']) {
        $u = Normalize-Url 'https://iranstatis.com/' $po.link
        if ($u) { Enqueue-Page $u $site 0 }
    }
    foreach ($m in $rest['media']) {
        $u = $m.source_url
        if (-not $u) { continue }
        $kind = Get-UrlKind $u
        if ($kind -eq 'page') { $kind = 'other' }
        $e = Get-Ext $u
        if ($e -in @('csv', 'xlsx', 'xls', 'dbf', 'shp', 'geojson', 'zip', 'json')) { $kind = 'data' }
        Enqueue-Asset $u $kind $site
    }

    # ---- 1b) Sitemap fallback (complete URL list if REST API was blocked)
    Write-Log 'دریافت فهرست صفحات از Sitemap (پشتیبان REST API)...'
    $sitemapLocs = New-Object System.Collections.Generic.List[string]
    foreach ($smBase in @('https://iranstatis.com/wp-sitemap.xml', 'https://iranstatis.com/sitemap_index.xml', 'https://iranstatis.com/sitemap.xml')) {
        try {
            $smb = Read-Api $smBase
            $smtxt = [Text.Encoding]::UTF8.GetString($smb)
            if ($smtxt -match '(?is)<sitemapindex.*?</sitemapindex>') {
                $children = [regex]::Matches($smtxt, '(?is)<loc>\s*(.*?)\s*</loc>') | ForEach-Object { $_.Groups[1].Value.Trim() }
                foreach ($cu in $children) {
                    try {
                        $cb = Read-Api $cu
                        $ct2 = [Text.Encoding]::UTF8.GetString($cb)
                        foreach ($cm in [regex]::Matches($ct2, '(?is)<loc>\s*(.*?)\s*</loc>')) { $sitemapLocs.Add($cm.Groups[1].Value.Trim()) }
                    } catch { Write-Log "خطا در زیرسایت‌نقشه $cu : $($_.Exception.Message)" 'WARN' }
                }
            }
            elseif ($smtxt -match '(?is)<urlset') {
                foreach ($cm in [regex]::Matches($smtxt, '(?is)<loc>\s*(.*?)\s*</loc>')) { $sitemapLocs.Add($cm.Groups[1].Value.Trim()) }
            }
            if ($sitemapLocs.Count -gt 0) { break }
        } catch { Write-Log "خطا در sitemap $smBase : $($_.Exception.Message)" 'WARN' }
    }
    if ($sitemapLocs.Count -gt 0) {
        $smCount = 0
        foreach ($su in ($sitemapLocs | Get-Unique)) {
            $u = Normalize-Url 'https://iranstatis.com/' $su
            if ($u) { Enqueue-Page $u $site 0; $smCount++ }
        }
        Write-Log "  sitemap: $smCount آدرس اضافه شد" 'DONE'
    }
    else {
        Write-Log '  sitemap در دسترس نبود (crawl از صفحه اصلی ادامه می‌یابد)' 'WARN'
    }

    # ---- 2) Data Lab API (آزمایشگاه داده)
    Write-Log 'دانلود داده‌های آزمایشگاه داده (API اختصاصی)...'
    $dl = Get-KindFolder $site 'apidata'
    if (-not (Test-Path $dl)) { New-Item -ItemType Directory -Path $dl -Force | Out-Null }
    try {
        $b = Read-Api 'https://iranstatis.com/wp-json/iranstatis/v1/sources'
        [IO.File]::WriteAllBytes((Join-Path $dl 'sources.json'), $b)
        $sources = $b | ConvertFrom-Json
        $publicKeys = @($sources | Where-Object { $_.access_level -eq 0 } | ForEach-Object { $_.key })
        $proKeys = @($sources | Where-Object { $_.access_level -eq 1 } | ForEach-Object { $_.key })
        Write-Log ("  منابع: {0} کل | {1} عمومی | {2} حرفه‌ای" -f $sources.Count, $publicKeys.Count, $proKeys.Count)
        $indDir = Join-Path $dl 'indicators'
        if (-not (Test-Path $indDir)) { New-Item -ItemType Directory -Path $indDir -Force | Out-Null }
        $okInd = 0
        foreach ($k in $publicKeys) {
            try {
                $kb = [Uri]::EscapeDataString($k)
                $ib = Read-Api "https://iranstatis.com/wp-json/iranstatis/v1/indicator?key=$kb"
                $ip = Join-Path $indDir "$k.json"
                [IO.File]::WriteAllBytes($ip, $ib)
                Save-Record "api:indicator:$k" (New-Record "https://iranstatis.com/wp-json/iranstatis/v1/indicator?key=$kb" 'apidata' $site 200 $ib.Length $ip '')
                Bump-Stat 'Done'; Bump-Stat 'Bytes' $ib.Length
                $rows = $ib | ConvertFrom-Json
                if ($rows -is [System.Array] -and $rows.Count -gt 0) {
                    $csvLines = @('"سال","استان","مقدار"')
                    foreach ($row in $rows) {
                        $g = ([string]$row.geo -replace '"', '""')
                        $csvLines += ('"{0}","{1}","{2}"' -f $row.year, $g, $row.value)
                    }
                    [System.IO.File]::WriteAllLines((Join-Path $indDir "$k.csv"), $csvLines, [System.Text.UTF8Encoding]::new($true))
                }
                $okInd++
            } catch { Write-Log "خطا در indicator $k : $($_.Exception.Message)" 'WARN' }
        }
        Write-Log "  آمارهای عمومی: $okInd از $($publicKeys.Count) دانلود شد" 'DONE'
        if ($ProToken) {
            foreach ($k in $proKeys) {
                try {
                    $kb = [Uri]::EscapeDataString($k)
                    $tk = [Uri]::EscapeDataString($ProToken)
                    $ib = Read-Api "https://iranstatis.com/wp-json/iranstatis/v1/indicator?key=$kb&token=$tk"
                    [IO.File]::WriteAllBytes((Join-Path $indDir "$k.json"), $ib)
                    Bump-Stat 'Done'
                } catch {
                    Save-Record "api:indicator:$k" (New-Record "https://iranstatis.com/wp-json/iranstatis/v1/indicator?key=$kb" 'apidata' $site 0 0 '' $($_.Exception.Message))
                    Bump-Stat 'Errors'
                }
            }
        }
        else {
            [IO.File]::WriteAllText((Join-Path $dl 'PRO-keys-not-downloaded.txt'), ($proKeys -join "`n"), [System.Text.UTF8Encoding]::new($true))
        }
    } catch { Write-Log "خطا در sources: $($_.Exception.Message)" 'ERROR' }

    foreach ($pair in @('https://iranstatis.com/wp-json/isde/v1/filters|isde-filters.json', 'https://iranstatis.com/wp-json/isde/v1/status|isde-status.json')) {
        $u3 = $pair.Split('|')[0]; $f3 = $pair.Split('|')[1]
        try {
            $b = Read-Api $u3
            [IO.File]::WriteAllBytes((Join-Path $dl $f3), $b)
            Save-Record "api:$f3" (New-Record $u3 'apidata' $site 200 $b.Length (Join-Path $dl $f3) '')
            Bump-Stat 'Done'
        } catch { Write-Log "خطا در $u3 : $($_.Exception.Message)" 'WARN' }
    }

    # seed homepage for BFS (catches theme assets + pages not in REST)
    $null = $script:PagesSeen.TryAdd('https://iranstatis.com/', 1)
    $script:PagesCount['https://iranstatis.com/'] = 1
    $script:WorkQueue.Enqueue((New-Item2 'https://iranstatis.com/' 'page' $site 0))

    Invoke-CrawlEngine $site $Workers

    Export-ManifestCsv (Join-Path (Join-Path $OutputRoot "$site\گزارش") 'manifest.csv') $site
    Write-SiteSummary $site
}
#endregion

#region --- Generic site crawl (imidro / behinyab / mimt / maj) ---
function Invoke-ExtraSite([string]$SiteKey) {
    $cfg = $script:ExtraSites[$SiteKey]
    $site = $cfg.Host
    Write-Log ("==== شروع crawl {0} ====" -f $site)
    $script:Stats = [pscustomobject]@{ Done = [long]0; Errors = [long]0; Bytes = [long]0; StatsLock = (New-Object object) }
    $script:WorkQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    $script:PagesSeen = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $script:AssetsSeen = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $script:PagesCount = [System.Collections.Concurrent.ConcurrentDictionary[string, object]]::new()
    $script:StopCrawl = $false
    Load-ExistingManifest $site

    foreach ($seed in $cfg.Seeds) {
        $seedAbs = $seed
        if ($seedAbs -notmatch '^https?://') { $seedAbs = 'https://' + $seedAbs }
        $kind = Get-UrlKind $seedAbs
        if ($kind -eq 'page') { Enqueue-Page $seedAbs $site 0 } else { Enqueue-Asset $seedAbs $kind $site }
    }

    Invoke-CrawlEngine $site $Workers

    Export-ManifestCsv (Join-Path (Join-Path $OutputRoot "$site\گزارش") 'manifest.csv') $site
    Write-SiteSummary $site
}
#endregion

#region --- Summary ---
function Get-DirSize([string]$Dir) {
    if (-not (Test-Path $Dir)) { return [long]0 }
    $s = [long]0
    Get-ChildItem -Path $Dir -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { $s += $_.Length }
    return $s
}

function Write-SiteSummary([string]$Site) {
    $base = Join-Path $OutputRoot $Site
    $lines = @()
    $lines += ('خلاصه دانلود {0}' -f $Site)
    $lines += (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    $lines += '--------------------------------------------------'
    foreach ($sub in @('صفحات', 'اسناد', 'تصاویر', 'چندرسانه', 'کدواستایل', 'فیدها', 'داده‌ها', 'داده-آزمایشگاه', 'سایر')) {
        $d = Join-Path $base $sub
        if (Test-Path $d) {
            $files = (Get-ChildItem -Path $d -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
            $size = Get-DirSize $d
            $lines += ('{0,-18} {1,8} فایل   {2,12:N0} بایت' -f $sub, $files, $size)
        }
    }
    $totalSize = Get-DirSize $base
    $totalFiles = (Get-ChildItem -Path $base -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
    $lines += '--------------------------------------------------'
    $lines += ('{0,-18} {1,8} فایل   {2,12:N0} بایت' -f 'مجموع', $totalFiles, $totalSize)
    $lines += ''
    $lines += ("انجام‌شده: {0} | خطا: {1} | حجم این اجرا: {2:N0} MB" -f $script:Stats.Done, $script:Stats.Errors, ($script:Stats.Bytes / 1MB))
    $txt = ($lines -join "`r`n")
    [IO.File]::WriteAllText((Join-Path (Join-Path $base 'گزارش') 'خلاصه.txt'), $txt, [System.Text.UTF8Encoding]::new($true))
    Write-Host $txt
    $csv = Join-Path (Join-Path $base 'گزارش') 'خطاها.csv'
    $lines2 = @('URL,HTTPStatus,Error,Type')
    foreach ($k in $script:Manifest.Keys) {
        $r = $script:Manifest[$k]
        if ($r.Site -eq $Site -and $r.HTTPStatus -ne 200) {
            $lines2 += ('"{0}",{1},"{2}","{3}"' -f $r.URL, $r.HTTPStatus, ($r.Error -replace '"', '""'), $r.Type)
        }
    }
    [System.IO.File]::WriteAllLines($csv, $lines2, [System.Text.UTF8Encoding]::new($true))
}
#endregion

#region --- Main ---
$swAll = [Diagnostics.Stopwatch]::StartNew()
Write-Log "==================================================="
Write-Log ("شروع دانلود | هدف: {0} | سایت: {1} | هم‌زمانی: {2} | نرخ: {3}/s" -f $OutputRoot, $Site, $Workers, $RatePerSec)
Write-Log "==================================================="

if ($Site -in @('both', 'all', 'amar')) {
    try { Invoke-AmarCrawl } catch { Write-Log ("خطای مهم در crawl amar.org.ir — ادامه با سایت‌های بعدی: " + $_.Exception.Message) 'ERROR' }
}
if ($Site -in @('both', 'all', 'iranstatis')) {
    try { Invoke-IranstatisCrawl } catch { Write-Log ("خطای مهم در crawl iranstatis.com — ادامه با سایت‌های بعدی: " + $_.Exception.Message) 'ERROR' }
}
foreach ($key in @('imidro', 'behinyab', 'mimt', 'maj')) {
    if ($Site -in @('all', $key)) {
        try { Invoke-ExtraSite $key } catch { Write-Log ("خطای مهم در crawl {0} — ادامه با سایت‌های بعدی: {1}" -f $script:ExtraSites[$key].Host, $_.Exception.Message) 'ERROR' }
    }
}

# overall summary
$allLines = @()
$allLines += 'خلاصه کلی دانلود محتوای سایت‌های آماری'
$allLines += (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
$allLines += '=================================================='
$grandSize = [long]0; $grandFiles = [long]0
$allHosts = @('amar.org.ir', 'iranstatis.com')
foreach ($k in $script:ExtraSites.Keys) { $allHosts += $script:ExtraSites[$k].Host }
foreach ($s2 in $allHosts) {
    $base = Join-Path $OutputRoot $s2
    if (Test-Path $base) {
        $sz = Get-DirSize $base
        $fc = (Get-ChildItem -Path $base -Recurse -File -ErrorAction SilentlyContinue | Measure-Object).Count
        $grandSize += $sz; $grandFiles += $fc
        $allLines += ('{0,-20} {1,8} فایل   {2,14:N0} بایت' -f $s2, $fc, $sz)
    }
}
$allLines += '=================================================='
$allLines += ('{0,-20} {1,8} فایل   {2,14:N0} بایت   ({3:N0} MB)' -f 'مجموع', $grandFiles, $grandSize, ($grandSize / 1MB))
$allLines += ("زمان کل: {0:N0} دقیقه" -f $swAll.Elapsed.TotalMinutes)
[IO.File]::WriteAllText((Join-Path $OutputRoot 'خلاصه-کلی.txt'), ($allLines -join "`r`n"), [System.Text.UTF8Encoding]::new($true))
Write-Host ($allLines -join "`r`n")
Write-Log "تمام شد. خلاصه: $OutputRoot\خلاصه-کلی.txt" 'DONE'
$script:Hc.Dispose()
#endregion
