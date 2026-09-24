# relaycheck-mullvad.ps1 - which Mullvad server should I use, and where is the problem?
#
# Works from any network: every run detects your location, ISP, home router and your ISP's first
# router, and picks the nearest Mullvad location in another country as the "international" check.
#
# Each run (about a minute):
#   1. Pings a ladder of targets 50 times each, all at once:
#        home network        your router (default gateway)
#        your isp            your isp's first router, 1.1.1.1 and 8.8.8.8 (anycast, served near you)
#        isp international   the 2 nearest Mullvad servers outside your country
#        region              every Mullvad server in the reference city (-RefCity, default los angeles)
#        city                every Mullvad server in the city under test (-City, default seattle), by provider
#   2. Gives every level a state from its share of lossy targets: clean (none), minor (up to 25%),
#      degraded (over 25%, under 60%), problem (60% or more). The headline is the first level from
#      the top that is a problem, else the first that is degraded, else any lossy city servers as minor.
#   3. Recommends the city server that is clean now and has the best 24h track record, plus a
#      backup from a different provider.
#   4. Speed test (speedtest.net, same server each time): direct, then through the recommended server.
#   5. Saves to relaycheck-mullvad.db and rebuilds relaycheck-mullvad.html.
#
# The laptop's Mullvad app is disconnected while pinging (so pings take the real route), connected to
# the recommended server only for the speed test, then put back the way it was.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File relaycheck-mullvad.ps1                          one run
#   powershell -ExecutionPolicy Bypass -File relaycheck-mullvad.ps1 -NoSpeed                 skip the speed test
#   powershell -ExecutionPolicy Bypass -File relaycheck-mullvad.ps1 -Loop 30                 run, wait 30 min, repeat
#   powershell -ExecutionPolicy Bypass -File relaycheck-mullvad.ps1 -Serve -Loop 30          dashboard on localhost:8765, run every 30 min
#   powershell -ExecutionPolicy Bypass -File relaycheck-mullvad.ps1 -City lax -RefCity sea   test another Mullvad city
#
# relaycheck-mullvad.db (SQLite; opens in DB Browser for SQLite):
#   runs     one row per run: where you were, diagnosis, recommended server, backup, speed
#   checks   one row per target per run: grp, name, provider, ip, loss, ms, jitter

param(
  [switch]$NoSpeed,
  [int]$Loop = 0,
  [int]$Pings = 50,
  [switch]$Serve,
  [int]$Port = 8765,
  [string]$City = "sea",
  [string]$RefCity = "lax",
  [string]$RegionName = "us-west-coast"
)

# ---------------------------------------------------------------- settings
# Loss thresholds (percent of pings lost): a target is clean at $CleanPct or less, lossy at $BadPct or
# more or when it gives no reply. A level is minor up to $GroupOk of its targets lossy, a problem from
# $GroupBad, degraded in between. A provider is only named when $BlameMin of its servers are lossy.
# $Ookla lists the preferred speedtest.net servers per city, so each run tests against the same one.
# Times are stored in the invariant culture, so the database reads the same on any Windows language.

$ErrorActionPreference = "Continue"
$Dir       = $PSScriptRoot
$DbPath    = Join-Path $Dir "relaycheck-mullvad.db"
$Report    = Join-Path $Dir "relaycheck-mullvad.html"
$CleanPct  = 2
$BadPct    = 4
$GroupOk   = 0.25
$GroupBad  = 0.6
$BlameMin  = 2
$SlowMbps  = 10
$IntlCount = 2
$HistHours = 24
$HistRows  = 24
$FlushDays = 7
$Ookla     = @{ sea = @(8864, 1782, 43122, 64151); lax = @(7190, 5905) }
$Dot       = [char]0x00B7
$Inv       = [Globalization.CultureInfo]::InvariantCulture
$TimeFmt   = "yyyy-MM-dd HH:mm:ss"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Net.ServicePointManager]::DefaultConnectionLimit = 64

function Write-Log($msg) { Write-Host ("[{0}] {1}" -f (Get-Date).ToString("HH:mm:ss", $Inv), $msg) }
function Get-Now { (Get-Date).ToString($TimeFmt, $Inv) }
function Get-Pct($a, $b) { if ($a -eq $null -or -not $b) { $null } else { [math]::Round($a / $b * 100) } }
function Test-Lossy($t) { $t.Ms -eq $null -or $t.Loss -ge $BadPct }
function Test-Clean($t) { $t.Ms -ne $null -and $t.Loss -le $CleanPct }
function Get-LostPings { [math]::Ceiling($BadPct * $Pings / 100) }

# ---------------------------------------------------------------- sqlite (winsqlite3.dll ships with Windows 10+)
if (-not ("RelaycheckMullvad.Sqlite" -as [type])) {
  Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace RelaycheckMullvad {
  public class Sqlite : IDisposable {
    const string L = "winsqlite3.dll";
    [DllImport(L)] static extern int sqlite3_open_v2(byte[] file, out IntPtr db, int flags, IntPtr vfs);
    [DllImport(L)] static extern int sqlite3_close_v2(IntPtr db);
    [DllImport(L)] static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int n, out IntPtr stmt, IntPtr tail);
    [DllImport(L)] static extern int sqlite3_step(IntPtr stmt);
    [DllImport(L)] static extern int sqlite3_finalize(IntPtr stmt);
    [DllImport(L)] static extern int sqlite3_bind_null(IntPtr stmt, int i);
    [DllImport(L)] static extern int sqlite3_bind_int64(IntPtr stmt, int i, long v);
    [DllImport(L)] static extern int sqlite3_bind_double(IntPtr stmt, int i, double v);
    [DllImport(L)] static extern int sqlite3_bind_text(IntPtr stmt, int i, byte[] v, int n, IntPtr destructor);
    [DllImport(L)] static extern int sqlite3_column_count(IntPtr stmt);
    [DllImport(L)] static extern IntPtr sqlite3_column_name(IntPtr stmt, int i);
    [DllImport(L)] static extern int sqlite3_column_type(IntPtr stmt, int i);
    [DllImport(L)] static extern long sqlite3_column_int64(IntPtr stmt, int i);
    [DllImport(L)] static extern double sqlite3_column_double(IntPtr stmt, int i);
    [DllImport(L)] static extern IntPtr sqlite3_column_text(IntPtr stmt, int i);
    [DllImport(L)] static extern int sqlite3_column_bytes(IntPtr stmt, int i);
    [DllImport(L)] static extern IntPtr sqlite3_errmsg(IntPtr db);

    static readonly IntPtr TRANSIENT = new IntPtr(-1);
    IntPtr db;

    static byte[] U(string s) { return Encoding.UTF8.GetBytes(s + "\0"); }

    static string S(IntPtr p, int n) {
      if (p == IntPtr.Zero) return null;
      byte[] b = new byte[n];
      Marshal.Copy(p, b, 0, n);
      return Encoding.UTF8.GetString(b);
    }

    static string Z(IntPtr p) {
      if (p == IntPtr.Zero) return null;
      int n = 0;
      while (Marshal.ReadByte(p, n) != 0) n++;
      return S(p, n);
    }

    string Err() { return Z(sqlite3_errmsg(db)); }

    public Sqlite(string path) {
      if (sqlite3_open_v2(U(path), out db, 0x2 | 0x4 | 0x10000, IntPtr.Zero) != 0) throw new Exception("open failed: " + Err());
    }

    IntPtr Prepare(string sql, object[] args) {
      IntPtr st;
      if (sqlite3_prepare_v2(db, U(sql), -1, out st, IntPtr.Zero) != 0) throw new Exception(Err() + " in: " + sql);
      if (args == null) return st;
      for (int i = 0; i < args.Length; i++) {
        object a = args[i];
        int k = i + 1;
        if (a == null || a is DBNull) sqlite3_bind_null(st, k);
        else if (a is bool) sqlite3_bind_int64(st, k, (bool)a ? 1 : 0);
        else if (a is int || a is long || a is short || a is byte) sqlite3_bind_int64(st, k, Convert.ToInt64(a));
        else if (a is double || a is float || a is decimal) sqlite3_bind_double(st, k, Convert.ToDouble(a));
        else {
          byte[] b = Encoding.UTF8.GetBytes(a.ToString());
          sqlite3_bind_text(st, k, b, b.Length, TRANSIENT);
        }
      }
      return st;
    }

    public void Exec(string sql, object[] args) {
      IntPtr st = Prepare(sql, args);
      try {
        int rc = sqlite3_step(st);
        if (rc != 101 && rc != 100) throw new Exception(Err() + " in: " + sql);
      } finally {
        sqlite3_finalize(st);
      }
    }

    public List<Dictionary<string, object>> Query(string sql, object[] args) {
      var rows = new List<Dictionary<string, object>>();
      IntPtr st = Prepare(sql, args);
      try {
        int n = sqlite3_column_count(st);
        while (true) {
          int rc = sqlite3_step(st);
          if (rc == 101) break;
          if (rc != 100) throw new Exception(Err() + " in: " + sql);
          var r = new Dictionary<string, object>();
          for (int i = 0; i < n; i++) {
            string name = Z(sqlite3_column_name(st, i));
            switch (sqlite3_column_type(st, i)) {
              case 1: r[name] = sqlite3_column_int64(st, i); break;
              case 2: r[name] = sqlite3_column_double(st, i); break;
              case 5: r[name] = null; break;
              default: r[name] = S(sqlite3_column_text(st, i), sqlite3_column_bytes(st, i)); break;
            }
          }
          rows.Add(r);
        }
      } finally {
        sqlite3_finalize(st);
      }
      return rows;
    }

    public void Dispose() {
      if (db == IntPtr.Zero) return;
      sqlite3_close_v2(db);
      db = IntPtr.Zero;
    }
  }
}
'@
}

function Get-Rows($db, [string]$sql, [object[]]$p = @()) {
  foreach ($r in $db.Query($sql, $p)) {
    $o = [ordered]@{}
    foreach ($k in $r.Keys) { $o[$k] = $r[$k] }
    [pscustomobject]$o
  }
}

function Get-Scalar($db, [string]$sql, [object[]]$p = @()) {
  $r = $db.Query($sql, $p)
  if ($r.Count) { @($r[0].Values)[0] }
}

function Add-Row($db, [string]$table, $h) {
  $cols = @($h.Keys)
  $vals = foreach ($c in $cols) { $h[$c] }
  $marks = ($cols | ForEach-Object { "?" }) -join ","
  $db.Exec("INSERT INTO $table (" + ($cols -join ",") + ") VALUES ($marks)", [object[]]@($vals))
  Get-Scalar $db "SELECT last_insert_rowid()"
}

$RunColumns = [ordered]@{
  time         = "TEXT NOT NULL UNIQUE"
  ip           = "TEXT"
  country      = "TEXT"
  city         = "TEXT"
  isp          = "TEXT"
  router_vpn   = "INTEGER"
  duration_s   = "REAL"
  level        = "TEXT"
  severity     = "TEXT"
  problem      = "TEXT"
  diagnosis    = "TEXT"
  pick         = "TEXT"
  pick_why     = "TEXT"
  backup       = "TEXT"
  speed_server = "TEXT"
  direct_down  = "REAL"
  direct_up    = "REAL"
  vpn_down     = "REAL"
  vpn_up       = "REAL"
  speed_note   = "TEXT"
}

function Open-Db {
  $db = New-Object RelaycheckMullvad.Sqlite $DbPath
  $db.Exec("PRAGMA journal_mode=WAL", $null)
  $db.Exec("PRAGMA foreign_keys=ON", $null)
  $db.Exec("PRAGMA busy_timeout=5000", $null)
  $cols = ($RunColumns.Keys | ForEach-Object { "$_ $($RunColumns[$_])" }) -join ", "
  $db.Exec("CREATE TABLE IF NOT EXISTS runs (id INTEGER PRIMARY KEY, $cols)", $null)
  $db.Exec("CREATE TABLE IF NOT EXISTS checks (run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
    grp TEXT, name TEXT, provider TEXT, ip TEXT, loss REAL, ms REAL, jitter REAL)", $null)
  $db.Exec("CREATE INDEX IF NOT EXISTS ix_checks ON checks(grp, name, run_id)", $null)

  # Databases made by older versions get any missing columns
  $have = @(Get-Rows $db "PRAGMA table_info(runs)" | ForEach-Object { $_.name })
  foreach ($c in $RunColumns.Keys) {
    if ($have -notcontains $c) { $db.Exec("ALTER TABLE runs ADD COLUMN $c " + ($RunColumns[$c] -replace " NOT NULL UNIQUE", ""), $null) }
  }
  $db
}

function Use-Db([scriptblock]$body) {
  $db = Open-Db
  try { & $body $db } finally { $db.Dispose() }
}

function Remove-OldRuns([int]$days) {
  $cut = (Get-Date).AddDays(-$days).ToString($TimeFmt, $Inv)
  Use-Db {
    param($db)
    $n = [int](Get-Scalar $db "SELECT COUNT(*) FROM runs WHERE time < ?" @($cut))
    $db.Exec("DELETE FROM runs WHERE time < ?", @($cut))
    [pscustomobject]@{ Removed = $n; Before = $cut }
  }
}

# ---------------------------------------------------------------- mullvad app
$Mullvad = @((Get-Command mullvad -ErrorAction SilentlyContinue).Source, "$env:ProgramFiles\Mullvad VPN\resources\mullvad.exe") |
  Where-Object { $_ -and (Test-Path $_) } |
  Select-Object -First 1

function Test-VpnUp {
  if (-not $Mullvad) { return $false }
  [bool](@(& $Mullvad status)[0] -match '^(Connected|Connecting)')
}

function Set-RelayHost($name) {
  $p = $name.Split("-")
  & $Mullvad relay set location $p[0] $p[1] $name | Out-Null
}

function Get-VpnState {
  if (-not $Mullvad) { return [pscustomobject]@{ Up = $false; Location = $null } }
  $loc = (@(& $Mullvad relay get) | Select-String '^\s*Location:') -replace '.*Location:\s*', ""
  [pscustomobject]@{ Up = (Test-VpnUp); Location = "$loc".Trim() }
}

function Restore-Vpn($s) {
  if (-not $Mullvad) { return }

  # Turns "city sea, us, hostname us-sea-wg-001", "city sea, us", "country my" or "any" back into relay set location
  switch -regex ($s.Location) {
    'hostname\s+(\S+)' { Set-RelayHost $matches[1]; break }
    'city\s+(\w+),\s*(\w+)' { & $Mullvad relay set location $matches[2] $matches[1] | Out-Null; break }
    'country\s+(\w+)' { & $Mullvad relay set location $matches[1] | Out-Null; break }
    '^any$' { & $Mullvad relay set location any | Out-Null; break }
    default { if ($s.Location) { Write-Log "couldn't restore mullvad location '$($s.Location)'; set it again in the app" } }
  }

  $now = Test-VpnUp
  if ($s.Up -and -not $now) { & $Mullvad connect --wait | Out-Null }
  elseif (-not $s.Up -and $now) { & $Mullvad disconnect --wait | Out-Null }
}

function Connect-Relay($name) {
  Set-RelayHost $name
  & $Mullvad connect --wait | Out-Null
  Start-Sleep 3
  [bool](@(& $Mullvad status)[0] -match '^Connected')
}

function Disconnect-Vpn {
  if ($Mullvad) { & $Mullvad disconnect --wait | Out-Null }
}

# ---------------------------------------------------------------- where am i (detected every run)
function Get-Relays {
  $app = Invoke-RestMethod "https://api.mullvad.net/app/v1/relays" -TimeoutSec 30
  foreach ($r in @($app.wireguard.relays)) {
    if (-not $r.active -or -not $r.ipv4_addr_in) { continue }
    $l = $app.locations.($r.location)
    $cityCode = $r.location.Split("-")[1]
    [pscustomobject]@{
      Name     = $r.hostname
      Provider = $r.provider
      IP       = $r.ipv4_addr_in
      Location = $r.location
      CityCode = $cityCode
      Country  = $l.country
      City     = ($l.city -replace ',.*$', "")
      Lat      = [double]$l.latitude
      Lon      = [double]$l.longitude
    }
  }
}

function Get-Km($lat1, $lon1, $lat2, $lon2) {
  $r = [math]::PI / 180
  $a = [math]::Sin(($lat2 - $lat1) * $r / 2)
  $b = [math]::Sin(($lon2 - $lon1) * $r / 2)
  6371 * 2 * [math]::Asin([math]::Sqrt($a * $a + [math]::Cos($lat1 * $r) * [math]::Cos($lat2 * $r) * $b * $b))
}

function Test-Private($ip) {
  $ip -match '^(10\.|127\.|169\.254\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.|100\.(6[4-9]|[7-9]\d|1[01]\d|12[0-7])\.)'
}

function Get-Here($relays) {
  try { $me = Invoke-RestMethod "https://am.i.mullvad.net/json" -TimeoutSec 15 } catch { $me = $null }
  $here = [pscustomobject]@{
    Ip         = $me.ip
    Country    = $me.country
    City       = $me.city
    Isp        = $me.organization
    Lat        = $me.latitude
    Lon        = $me.longitude
    ViaMullvad = [bool]($me -and $me.mullvad_exit_ip)
    Gateway    = $null
    IspHop     = $null
    Intl       = @()
    IntlPlace  = $null
  }

  # Home router: the lowest-metric default route, whatever adapter it is on
  $route = Get-NetRoute -DestinationPrefix 0.0.0.0/0 -ErrorAction SilentlyContinue | Sort-Object RouteMetric, InterfaceMetric | Select-Object -First 1
  $here.Gateway = $route.NextHop

  # Your isp's first router: the first public hop on the way out, past home routers, modems and carrier-grade nat
  $hops = @(tracert -d -h 8 -w 700 1.1.1.1 | ForEach-Object { if ($_ -match '^\s*\d+\s.*?(\d+\.\d+\.\d+\.\d+)\s*$') { $matches[1] } })
  $here.IspHop = $hops | Where-Object { -not (Test-Private $_) -and $_ -ne "1.1.1.1" } | Select-Object -First 1

  # Isp international: the nearest mullvad location outside your country (mullvad writes USA and UK, the geo service full names)
  if ($me -and $me.latitude -ne $null -and $relays) {
    $alias = @{ "United States" = "USA"; "United Kingdom" = "UK" }
    $mine = if ($alias[$here.Country]) { $alias[$here.Country] } else { $here.Country }
    $locations = $relays | Group-Object Location | ForEach-Object {
      $f = $_.Group[0]
      [pscustomobject]@{ Country = $f.Country; City = $f.City; Km = (Get-Km $here.Lat $here.Lon $f.Lat $f.Lon); Relays = $_.Group }
    }
    $abroad = $locations | Where-Object { $_.Country -ne $mine -and $_.Km -gt 150 } | Sort-Object Km | Select-Object -First 1
    if ($abroad) {
      $here.Intl = @($abroad.Relays | Select-Object -First $IntlCount)
      $here.IntlPlace = "$($abroad.City), $($abroad.Country)"
    }
  }
  $here
}

# ---------------------------------------------------------------- ping every target at once
function Get-Targets($here, $cityRelays, $refRelays, $cityName, $refName) {
  $t = @()
  if ($here.Gateway) { $t += @{ Grp = "home network"; Name = "router $($here.Gateway)"; Provider = ""; IP = $here.Gateway; Infra = $true } }
  if ($here.IspHop) { $t += @{ Grp = "your isp"; Name = "isp router $($here.IspHop)"; Provider = $here.Isp; IP = $here.IspHop; Infra = $true } }
  $t += @{ Grp = "your isp"; Name = "1.1.1.1"; Provider = "cloudflare"; IP = "1.1.1.1"; Infra = $false }
  $t += @{ Grp = "your isp"; Name = "8.8.8.8"; Provider = "google"; IP = "8.8.8.8"; Infra = $false }
  $t += @($here.Intl | ForEach-Object { @{ Grp = "isp international"; Name = $_.Name; Provider = $_.Provider; IP = $_.IP; Infra = $false } })
  $t += @($refRelays | ForEach-Object { @{ Grp = $refName; Name = $_.Name; Provider = $_.Provider; IP = $_.IP; Infra = $false } })
  $t += @($cityRelays | ForEach-Object { @{ Grp = $cityName; Name = $_.Name; Provider = $_.Provider; IP = $_.IP; Infra = $false } })
  foreach ($x in $t) {
    $x.Count = $Pings
    $x.Loss = $null
    $x.Ms = $null
    $x.Jitter = $null
    $x.Ignored = $false
  }
  $t
}

function Test-Targets($targets) {
  $block = {
    param($t)
    $p = New-Object Net.NetworkInformation.Ping
    $ok = 0
    $ms = @()
    for ($i = 0; $i -lt $t.Count; $i++) {
      try {
        $x = $p.Send($t.IP, 1000)
        if ($x.Status -eq "Success") {
          $ok++
          $ms += $x.RoundtripTime
        }
      } catch {}
      Start-Sleep -Milliseconds 50
    }
    $jit = $null
    if ($ms.Count -gt 1) { $jit = ($(for ($i = 1; $i -lt $ms.Count; $i++) { [math]::Abs($ms[$i] - $ms[$i - 1]) }) | Measure-Object -Average).Average }
    $t.Loss = [math]::Round(($t.Count - $ok) / $t.Count * 100, 1)
    $t.Ms = if ($ms.Count) { [math]::Round(($ms | Measure-Object -Average).Average) } else { $null }
    $t.Jitter = if ($jit -ne $null) { [math]::Round($jit, 1) } else { $null }
    [pscustomobject]$t
  }

  $pool = [runspacefactory]::CreateRunspacePool(1, 64)
  $pool.Open()
  $jobs = foreach ($t in $targets) {
    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($block).AddArgument($t)
    [pscustomobject]@{ PS = $ps; H = $ps.BeginInvoke() }
  }
  $out = foreach ($j in $jobs) {
    $j.PS.EndInvoke($j.H)
    $j.PS.Dispose()
  }
  $pool.Close()
  $pool.Dispose()

  # A router that never answers ping (some block it) says nothing about loss, so leave it out when later levels answer
  $answering = @($out | Where-Object { -not $_.Infra -and $_.Ms -ne $null }).Count
  foreach ($r in $out) {
    if ($r.Infra -and $r.Ms -eq $null -and $answering) { $r.Ignored = $true }
  }
  $out
}

# ---------------------------------------------------------------- speedtest.net
function Select-OoklaServer($cityName, $cityCode) {
  $url = "https://www.speedtest.net/api/js/servers?engine=js&limit=30&search=" + [uri]::EscapeDataString($cityName)
  try { $raw = Invoke-RestMethod $url -TimeoutSec 20 } catch { return $null }

  # Assign before filtering: PS 5.1 pipes a JSON array as one object
  $list = @($raw | Where-Object { $_.name -like "$cityName*" })
  $pref = @($Ookla[$cityCode])
  $ordered = @($list | Where-Object { $pref -contains [int]$_.id } | Sort-Object { $pref.IndexOf([int]$_.id) }) +
    @($list | Where-Object { $pref -notcontains [int]$_.id })

  # 4 s probe: a server under 2 mbps on one connection is that server's problem, not the route's
  foreach ($s in $ordered | Select-Object -First 8) {
    $probe = "http://$($s.host)/download?nocache=probe&size=25000000"
    $r = "$(curl.exe -s -L -o NUL -m 4 -A "Mozilla/5.0 relaycheck-mullvad" -w "%{http_code} %{speed_download}" $probe)".Split(" ")
    if ($r[0] -eq "200" -and [double]$r[1] * 8 / 1e6 -ge 2) { return [pscustomobject]@{ Name = "$($s.sponsor) #$($s.id)"; Host = $s.host } }
  }
  $null
}

function Measure-Speed($hostName, $mode, [int]$seconds) {
  # Ookla servers answer 500 without a User-Agent
  $worker = {
    param($a)
    $buf = New-Object byte[] 262144
    if ($a.Mode -eq "up") { (New-Object Random).NextBytes($buf) }
    while ([DateTime]::UtcNow -lt $a.Deadline) {
      try {
        $nc = [guid]::NewGuid().ToString("N")
        if ($a.Mode -eq "down") {
          $req = [Net.HttpWebRequest]::Create("http://$($a.Host)/download?nocache=$nc&size=25000000")
          $req.UserAgent = "Mozilla/5.0 relaycheck-mullvad"
          $req.Timeout = 10000
          $req.ReadWriteTimeout = 10000
          $resp = $req.GetResponse()
          $s = $resp.GetResponseStream()
          while ([DateTime]::UtcNow -lt $a.Deadline) {
            $n = $s.Read($buf, 0, $buf.Length)
            if ($n -le 0) { break }
            $a.Counter[$a.Slot] += $n
          }
          $resp.Close()
        } else {
          $len = 4MB
          $req = [Net.HttpWebRequest]::Create("http://$($a.Host)/upload?nocache=$nc")
          $req.UserAgent = "Mozilla/5.0 relaycheck-mullvad"
          $req.Method = "POST"
          $req.ContentType = "application/octet-stream"
          $req.ContentLength = $len
          $req.AllowWriteStreamBuffering = $false
          $req.Timeout = 10000
          $req.ReadWriteTimeout = 10000
          $s = $req.GetRequestStream()
          $sent = 0
          while ($sent -lt $len) {
            $n = [math]::Min($buf.Length, $len - $sent)
            $s.Write($buf, 0, $n)
            $sent += $n
            $a.Counter[$a.Slot] += $n
            if ([DateTime]::UtcNow -ge $a.Deadline) { break }
          }
          if ($sent -ge $len) {
            $s.Close()
            $req.GetResponse().Close()
          } else {
            $req.Abort()
          }
        }
      } catch {
        Start-Sleep -Milliseconds 200
      }
    }
  }

  $conns = 8
  $counter = New-Object "long[]" $conns
  $deadline = [DateTime]::UtcNow.AddSeconds($seconds)
  $pool = [runspacefactory]::CreateRunspacePool(1, $conns)
  $pool.Open()
  $jobs = for ($i = 0; $i -lt $conns; $i++) {
    $ps = [powershell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($worker).AddArgument(@{ Host = $hostName; Mode = $mode; Deadline = $deadline; Counter = $counter; Slot = $i })
    [pscustomobject]@{ PS = $ps; H = $ps.BeginInvoke() }
  }

  # Measure from 2 s in, after tcp has ramped up
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $t0 = $null
  while ([DateTime]::UtcNow -lt $deadline) {
    Start-Sleep -Milliseconds 250
    if (-not $t0 -and $sw.Elapsed.TotalSeconds -ge 2) { $t0 = @{ T = $sw.Elapsed.TotalSeconds; B = ($counter | Measure-Object -Sum).Sum } }
  }
  $t1 = @{ T = $sw.Elapsed.TotalSeconds; B = ($counter | Measure-Object -Sum).Sum }

  foreach ($j in $jobs) {
    [void]$j.H.AsyncWaitHandle.WaitOne(12000)
    try { $j.PS.Stop() } catch {}
    $j.PS.Dispose()
  }
  $pool.Close()
  $pool.Dispose()

  if (-not $t1.B -or -not $t0) { return $null }
  [math]::Round(($t1.B - $t0.B) * 8 / ($t1.T - $t0.T) / 1e6, 1)
}

function Test-Speed($server) {
  [pscustomobject]@{ Down = (Measure-Speed $server.Host "down" 10); Up = (Measure-Speed $server.Host "up" 8) }
}

function Get-SpeedNotes($d, $v, $pick) {
  $dn = if (-not $v -or $v.Down -eq $null) { "couldn't test the vpn via $($pick.Name)." }
    elseif ($d.Down -eq $null) { "couldn't test direct download, so there's nothing to compare the vpn with." }
    elseif ($d.Down -lt $SlowMbps -and $v.Down -lt $SlowMbps) { "this city is slow even without the vpn (direct download $($d.Down) mbps), so the route is the bottleneck, not mullvad." }
    elseif ($d.Down -lt $SlowMbps) { "direct download was unusually slow this time ($($d.Down) mbps), so there's no fair comparison; the vpn got $($v.Down) mbps." }
    elseif ($v.Down -lt $SlowMbps) { "$($pick.Name) pings clean but the vpn is slow (download $($v.Down) mbps). try the backup server." }
    else { "the vpn is getting $(Get-Pct $v.Down $d.Down)% of the direct download speed." }

  $up = $null
  if ($v -and $v.Up -ne $null -and $d.Up) {
    $up = if ($v.Up -lt $SlowMbps -and $d.Up -ge $SlowMbps) { "upload through the vpn is slow (upload $($v.Up) mbps, $(Get-Pct $v.Up $d.Up)% of the direct upload speed)." }
      elseif ($d.Up -lt $SlowMbps -and $v.Up -ge $SlowMbps) { "direct upload was unusually slow this time ($($d.Up) mbps), so there's no fair comparison; the vpn got $($v.Up) mbps." }
      elseif ($d.Up -lt $SlowMbps) { "upload is slow both direct ($($d.Up) mbps) and through the vpn ($($v.Up) mbps)." }
      else { "the vpn is getting $(Get-Pct $v.Up $d.Up)% of the direct upload speed." }
  }
  @($dn, $up) | Where-Object { $_ }
}

# ---------------------------------------------------------------- diagnosis
# Every level gets one of four states from how many of its targets are lossy:
#   clean: none   minor: up to $GroupOk   degraded: in between   problem: $GroupBad or more
$Rank = @{ "n/a" = -1; clean = 0; minor = 1; degraded = 2; problem = 3 }

function Get-Ranges([int]$n) {
  # The lossy-target counts that give each state for a level of n targets, e.g. 11 gives 0, 1-2, 3-6, 7-11
  $minorMax = [math]::Floor($GroupOk * $n)
  $probMin = [math]::Max(1, [math]::Ceiling($GroupBad * $n))
  $span = { param($lo, $hi) if ($lo -gt $hi) { $null } elseif ($lo -eq $hi) { "$lo" } else { "$lo&ndash;$hi" } }
  [pscustomobject]@{ Clean = "0"; Minor = (& $span 1 $minorMax); Degraded = (& $span ($minorMax + 1) ($probMin - 1)); Problem = (& $span $probMin $n) }
}

function Get-Status($items) {
  # An empty function result arrives as $null, so drop nulls before counting
  $items = @($items | Where-Object { $_ -and -not $_.Ignored })
  $n = $items.Count
  if (-not $n) { return [pscustomobject]@{ Status = "n/a"; Bad = 0; N = 0 } }
  $bad = @($items | Where-Object { Test-Lossy $_ }).Count
  $share = $bad / $n
  $s = if ($bad -eq 0) { "clean" } elseif ($share -ge $GroupBad) { "problem" } elseif ($share -le $GroupOk) { "minor" } else { "degraded" }
  [pscustomobject]@{ Status = $s; Bad = $bad; N = $n }
}

function Select-Group($res, $grp) { @($res | Where-Object { $_.Grp -eq $grp }) }

function Format-Result($t) {
  if ($t.Ms -eq $null) { "no reply" } else { "loss $($t.Loss)% $Dot latency $($t.Ms) ms" }
}

function Format-Lossy($list) {
  $items = @($list | Where-Object { $_ }) | Sort-Object Name
  ($items | ForEach-Object { "$($_.Name) (" + $(if ($_.Ms -eq $null) { "no reply" } else { "loss $($_.Loss)%" }) + ")" }) -join ", "
}

function Format-Short($list) {
  # The first name in full, the rest by number, e.g. "us-sea-wg-404, 408, 409"
  $n = @(@($list | Where-Object { $_ }) | Sort-Object Name | ForEach-Object { $_.Name })
  if ($n.Count) { (@($n[0]) + @($n | Select-Object -Skip 1 | ForEach-Object { $_ -replace '^.*-wg-', "" })) -join ", " }
}

function Get-Diagnosis($res, $here, $cityName, $refName) {
  $hm = Get-Status (Select-Group $res "home network")
  $isp = Get-Status (Select-Group $res "your isp")
  $intl = Get-Status (Select-Group $res "isp international")
  $sea = @(Select-Group $res $cityName)
  $la = @(Select-Group $res $refName)
  $seaS = Get-Status $sea
  $laS = Get-Status $la
  $prov = @(foreach ($p in @($sea | Group-Object Provider)) {
    $s = Get-Status $p.Group
    [pscustomobject]@{ Provider = $p.Name; Status = $s.Status; Bad = $s.Bad; N = $s.N }
  })
  $lossy = @($sea | Where-Object { Test-Lossy $_ })
  $where = if ($here.Country) { $here.Country.ToLower() } else { "your country" }

  $provLine = { param($ps) (@($ps) | ForEach-Object { $pv = $_.Provider; "provider ${pv}: $($_.Bad) of $($_.N) lossy ($(Format-Lossy @($lossy | Where-Object { $_.Provider -eq $pv })))" }) -join " $Dot " }
  $okLine = { param($ps) (@($ps) | ForEach-Object { "provider $($_.Provider): $($_.Bad) of $($_.N) lossy" }) -join " $Dot " }
  $grpLine = { param($g) Format-Lossy @(Select-Group $res $g | Where-Object { -not $_.Ignored -and (Test-Lossy $_) }) }
  $common = @{ Home = $hm; Isp = $isp; Intl = $intl; Sea = $seaS; La = $laS; Providers = $prov }

  if (-not $seaS.N) {
    return [pscustomobject](@{ Level = "none"; Severity = "n/a"; Label = "no $cityName servers tested"; Text = "no active mullvad $cityName servers were found to test."; Blamed = @() } + $common)
  }

  # Problems first, top to bottom, then degraded levels, then any lossy servers on their own as minor
  $pick = $null
  $blamed = @()
  foreach ($sev in "problem", "degraded") {
    # A provider is named only when 2+ of its servers are lossy and another provider is healthier than this state
    $bp = @($prov | Where-Object { $_.Status -eq $sev -and $_.Bad -ge $BlameMin })
    $op = @($prov | Where-Object { $Rank[$_.Status] -lt $Rank[$sev] })
    $provPick = if ($bp.Count -and $op.Count) { "provider", "provider $(($bp.Provider) -join ', ')", "$(& $provLine $bp). $(& $okLine $op)." }
    $pick = if ($hm.Status -eq $sev) {
        "home", "home network", "home network: $($hm.Bad) of $($hm.N) lossy ($(& $grpLine "home network")). your wi-fi or router is dropping packets; fix this first."
      } elseif ($isp.Status -eq $sev) {
        "isp", "your isp", "your isp: $($isp.Bad) of $($isp.N) lossy ($(& $grpLine "your isp")). sites near you lose packets, so it has nothing to do with mullvad or $cityName."
      } elseif ($intl.Status -eq $sev) {
        "isp-intl", "isp international", "isp international: $($intl.Bad) of $($intl.N) lossy ($(& $grpLine "isp international")). your isp's links out of $where lose packets, so all traffic leaving $where is affected."
      } elseif ($seaS.Status -eq $sev -and $Rank[$laS.Status] -ge $Rank[$sev]) {
        if ($provPick) { $provPick } else { "region", $RegionName, "${RegionName}: $cityName $($seaS.Bad) of $($seaS.N) lossy, $refName $($laS.Bad) of $($laS.N) lossy. it's the route from $where, not mullvad." }
      } elseif ($seaS.Status -eq $sev) {
        if ($provPick) { $provPick } else { "city", $cityName, "${cityName}: $($seaS.Bad) of $($seaS.N) lossy across providers ($(Format-Lossy $lossy)). ${refName}: $($laS.Bad) of $($laS.N) lossy." }
      } elseif ($provPick) {
        $provPick
      }
    if ($pick) {
      $severity = $sev
      if ($pick[0] -eq "provider") { $blamed = @($bp.Provider) }
      break
    }
  }
  if (-not $pick) {
    if ($lossy.Count) {
      $severity = "minor"
      $pick = "server", (Format-Short $lossy), "${cityName}: $($lossy.Count) of $($sea.Count) lossy ($(Format-Lossy $lossy)), with no wider problem. the other $($sea.Count - $lossy.Count) $cityName servers are clean."
    } else {
      $severity = "clean"
      $pick = "none", "", "${cityName}: 0 of $($sea.Count) lossy."
    }
  }

  $level, $what, $text = $pick
  $label = if ($what) { "${severity}: $what" } else { $severity }
  [pscustomobject](@{ Level = $level; Severity = $severity; Label = $label; Text = $text; Blamed = $blamed } + $common)
}

# ---------------------------------------------------------------- recommended server
function Format-Server($r) {
  # One summary, the same wording everywhere: provider, loss, latency and track record, joined by middle dots
  if (-not $r) { return "none available" }
  "provider $($r.Provider) $Dot $(Format-Result $r) $Dot clean $($r.Clean) of $($r.Runs) runs (${HistHours}h)"
}

function Get-History($db, $grp, [datetime]$since) {
  # Saved runs and clean runs per server since a time; runs through a router tunnel don't count
  $sql = "SELECT c.name, COUNT(*) AS n, SUM(CASE WHEN c.ms IS NOT NULL AND c.loss <= ? THEN 1 ELSE 0 END) AS clean
          FROM checks c JOIN runs r ON r.id = c.run_id
          WHERE c.grp = ? AND r.time >= ? AND COALESCE(r.router_vpn, 0) = 0
          GROUP BY c.name"
  $hist = @{}
  Get-Rows $db $sql @($CleanPct, $grp, $since.ToString($TimeFmt, $Inv)) | ForEach-Object { $hist[$_.name] = $_ }
  $hist
}

function Get-Ranked($db, $grp, $list) {
  # Track record over the last $HistHours and the last $FlushDays (the same window Flush keeps), each plus this run, which isn't saved yet
  $day = Get-History $db $grp (Get-Date).AddHours(-$HistHours)
  $week = Get-History $db $grp (Get-Date).AddDays(-$FlushDays)

  $rows = foreach ($s in @($list | Where-Object { $_ })) {
    $now = if (Test-Clean $s) { 1 } else { 0 }
    $h = $day[$s.Name]
    $w = $week[$s.Name]
    $n = 1 + $(if ($h) { [int]$h.n } else { 0 })
    $c = $now + $(if ($h) { [int]$h.clean } else { 0 })
    $n7 = 1 + $(if ($w) { [int]$w.n } else { 0 })
    $c7 = $now + $(if ($w) { [int]$w.clean } else { 0 })
    [pscustomobject]@{
      Name     = $s.Name
      Provider = $s.Provider
      Loss     = $s.Loss
      Ms       = $s.Ms
      Jitter   = $s.Jitter
      Runs     = $n
      Clean    = $c
      Rate     = $c / $n
      Runs7    = $n7
      Clean7   = $c7
      Rate7    = $c7 / $n7
      Score    = ($c + 1) / ($n + 2)
      CleanNow = (Test-Clean $s)
    }
  }

  # Clean now first, then track record (smoothed, so 1 clean run doesn't beat 47 of 48), then loss and latency
  $byLoss = { if ($_.Ms -eq $null) { 999 } else { $_.Loss } }
  $byMs = { if ($_.Ms -eq $null) { 9999 } else { $_.Ms } }
  @($rows | Sort-Object @{ e = { -[int]$_.CleanNow } }, @{ e = { -$_.Score } }, @{ e = $byLoss }, @{ e = $byMs })
}

function Get-Pick($db, $cityRes, $cityName) {
  $ranked = @(Get-Ranked $db $cityName $cityRes)
  $pick = $ranked | Select-Object -First 1
  $backup = $ranked | Select-Object -Skip 1 | Where-Object { $_.Provider -ne $pick.Provider } | Select-Object -First 1
  if (-not $backup) { $backup = $ranked | Select-Object -Skip 1 -First 1 }
  $why = Format-Server $pick
  if ($pick -and -not $pick.CleanNow) { $why = "no $cityName server is clean right now; this one has the least loss ($why)" }
  [pscustomobject]@{ Pick = $pick; Backup = $backup; Why = $why; Ranked = $ranked }
}

# ---------------------------------------------------------------- one run
function Invoke-Run {
  $stamp = Get-Now
  $clock = [Diagnostics.Stopwatch]::StartNew()
  Write-Log "relaycheck-mullvad run"
  $vpn0 = Get-VpnState
  $run = $null
  try {
    if ($vpn0.Up) {
      Write-Log "disconnecting the laptop vpn while pinging"
      Disconnect-Vpn
      Start-Sleep 2
    }
    $relays = @(Get-Relays)
    if (-not $relays.Count) {
      Write-Log "couldn't load the mullvad server list; skipping this run"
      return
    }
    $cityRelays = @($relays | Where-Object { $_.CityCode -eq $City })
    $refRelays = @($relays | Where-Object { $_.CityCode -eq $RefCity })
    $cityName = if ($cityRelays.Count) { $cityRelays[0].City.ToLower() } else { $City }
    $refName = if ($refRelays.Count) { $refRelays[0].City.ToLower() } else { $RefCity }
    $here = Get-Here $relays

    $targets = @(Get-Targets $here $cityRelays $refRelays $cityName $refName)
    Write-Log "pinging $($targets.Count) targets $Pings times each"
    $res = @(Test-Targets $targets)
    $diag = Get-Diagnosis $res $here $cityName $refName
    $ranks = Use-Db {
      param($db)
      [pscustomobject]@{ Pick = (Get-Pick $db (Select-Group $res $cityName) $cityName); Ref = @(Get-Ranked $db $refName (Select-Group $res $refName)) }
    }
    $pick = $ranks.Pick

    $speed = $null
    if (-not $here.ViaMullvad -and -not $NoSpeed -and $pick.Pick) {
      $srv = Select-OoklaServer $cityName $City
      if ($srv) {
        Write-Log "speed test direct (no vpn) to $($srv.Name)"
        $d = Test-Speed $srv
        $v = $null
        if ($Mullvad) {
          Write-Log "speed test vpn via $($pick.Pick.Name)"
          if (Connect-Relay $pick.Pick.Name) { $v = Test-Speed $srv }
          Disconnect-Vpn
        }
        $speed = [pscustomobject]@{ Server = $srv.Name; Direct = $d; Vpn = $v; Notes = @(Get-SpeedNotes $d $v $pick.Pick) }
      }
    }

    $run = [pscustomobject]@{
      Time      = $stamp
      Here      = $here
      CityName  = $cityName
      RefName   = $refName
      Results   = $res
      Diag      = $diag
      Pick      = $pick
      RefRanked = $ranks.Ref
      Speed     = $speed
      Duration  = 0
      Pings     = $Pings
    }
  } catch {
    Write-Log "run failed: $($_.Exception.Message)"
    $run = $null
  } finally {
    Restore-Vpn $vpn0
  }
  if ($run) { $run.Duration = [math]::Round($clock.Elapsed.TotalSeconds) }
  $run
}

function Save-Run($run) {
  $h = $run.Here
  $s = $run.Speed
  $row = [ordered]@{
    time         = $run.Time
    ip           = $h.Ip
    country      = $h.Country
    city         = $h.City
    isp          = $h.Isp
    router_vpn   = $h.ViaMullvad
    duration_s   = $run.Duration
    level        = $run.Diag.Level
    severity     = $run.Diag.Severity
    problem      = $run.Diag.Label
    diagnosis    = $run.Diag.Text
    pick         = $run.Pick.Pick.Name
    pick_why     = $run.Pick.Why
    backup       = $run.Pick.Backup.Name
    speed_server = $s.Server
    direct_down  = $s.Direct.Down
    direct_up    = $s.Direct.Up
    vpn_down     = $s.Vpn.Down
    vpn_up       = $s.Vpn.Up
    speed_note   = $(if ($s) { $s.Notes -join " " })
  }
  Use-Db {
    param($db)
    $db.Exec("BEGIN", $null)
    try {
      $id = Add-Row $db "runs" $row
      foreach ($r in $run.Results) {
        [void](Add-Row $db "checks" ([ordered]@{ run_id = $id; grp = $r.Grp; name = $r.Name; provider = $r.Provider; ip = $r.IP; loss = $r.Loss; ms = $r.Ms; jitter = $r.Jitter }))
      }
      $db.Exec("COMMIT", $null)
    } catch {
      try { $db.Exec("ROLLBACK", $null) } catch {}
      Write-Log "database write failed: $($_.Exception.Message)"
    }
  }
}

# ---------------------------------------------------------------- report
# Headings, table headers and buttons are Title Case; all other page text is lower case (Format-Html lowercases it)
function Format-Html($s) { [System.Net.WebUtility]::HtmlEncode(([string]$s).ToLower()) }

function Format-Mbps($v) {
  if ($v -eq $null) { "&ndash;" } else { "{0} mbps" -f ([double]$v).ToString("0.0", $Inv) }
}

function Format-Ms($v, [switch]$Decimal) {
  if ($v -eq $null) { "&ndash;" } elseif ($Decimal) { "{0} ms" -f ([double]$v).ToString("0.0", $Inv) } else { "$v ms" }
}

function Get-Tone($loss, $ms) {
  if ($ms -eq $null) { "problem" } elseif ($loss -le $CleanPct) { "clean" } elseif ($loss -lt 10) { "degraded" } else { "problem" }
}

function Get-RateTone($rate) {
  if ($rate -ge 0.75) { "clean" } elseif ($rate -ge 0.4) { "degraded" } else { "problem" }
}

function Format-ServerTable($title, $rows, $pickName) {
  $h = "<section class=midsec><h2>$title</h2><div class=wrap><table><tr><th>Server</th><th>Provider</th><th class=n>Loss</th><th class=n>Latency</th><th class=n>Jitter</th><th>Clean (${HistHours}h)</th><th>Clean (${FlushDays}d)</th></tr>"
  foreach ($r in @($rows | Where-Object { $_ })) {
    # A bar full at 20% loss
    $w = if ($r.Ms -eq $null) { 100 } else { [math]::Min(100, $r.Loss * 5) }
    $lossTxt = if ($r.Ms -ne $null) { "$($r.Loss)%" } else { "no reply" }
    $me = if ($r.Name -eq $pickName) { " class=me" } else { "" }
    $h += "<tr$me><td class=nm>$(Format-Html $r.Name)</td><td class=mut>$(Format-Html $r.Provider)</td>" +
      "<td class=n><span class=bv><span class='bar $(Get-Tone $r.Loss $r.Ms)'><i style='width:$w%'></i></span><span class=v>$lossTxt</span></span></td>" +
      "<td class=n>$(Format-Ms $r.Ms)</td><td class=n>$(Format-Ms $r.Jitter -Decimal)</td>" +
      "<td><span class=bv><span class='bar wide $(Get-RateTone $r.Rate)'><i style='width:$([math]::Round($r.Rate * 100))%'></i></span><span class=mut>$($r.Clean) of $($r.Runs) runs</span></span></td>" +
      "<td><span class=bv><span class='bar wide $(Get-RateTone $r.Rate7)'><i style='width:$([math]::Round($r.Rate7 * 100))%'></i></span><span class=mut>$($r.Clean7) of $($r.Runs7) runs</span></span></td></tr>"
  }
  $h + "</table></div></section>"
}

function Format-Ladder($run) {
  $d = $run.Diag
  $res = $run.Results
  $h = $run.Here
  $cityRes = @(Select-Group $res $run.CityName)

  $step = {
    param($label, $s, $detail, $lvl, $tested, $n, $defText, [switch]$Hit)
    if (-not $defText -and $s -in "minor", "degraded", "problem" -and $n) {
      $r = (Get-Ranges $n).$s
      if ($r) { $defText = "$s = $r of $n lossy" }
    }
    $hitClass = if ($Hit -or $d.Level -eq $lvl) { " hit" } else { "" }
    $defHtml = if ($defText) { "<span class=def>$defText</span>" } else { "" }
    "<div class='step $s$hitClass'><i class=node></i><span class=lb>$(Format-Html $label)<small>$tested</small></span>" +
      "<span class=st><b class='$s'>$(Format-Html $s)</b><span class=mut>$(Format-Html $detail)</span>$defHtml</span></div>"
  }

  # The tested line lists every target with its own result, or "ignores ping" for a router that never answers
  $each = {
    param($g)
    $lines = Select-Group $res $g | ForEach-Object {
      $prov = if ($_.Provider -and $_.Name -notmatch '-wg-') { " ($(Format-Html $_.Provider))" } else { "" }
      $result = if ($_.Ignored) { "ignores ping, not counted" } else { Format-Html (Format-Result $_) }
      "$(Format-Html $_.Name)$prov &middot; $result"
    }
    $lines -join "<br>"
  }
  $cnt = { param($s) if ($s.N) { "$($s.Bad) of $($s.N) lossy" } else { "nothing to test" } }

  $homeTested = if ($h.Gateway) { & $each "home network" } else { "no default gateway found" }
  $intlTested = if ($h.IntlPlace) { "nearest mullvad location abroad: $(Format-Html $h.IntlPlace)<br>" + (& $each "isp international") } else { "couldn't work out your location" }
  $out = @(
    (& $step "home network" $d.Home.Status (& $cnt $d.Home) "home" $homeTested $d.Home.N),
    (& $step "your isp" $d.Isp.Status (& $cnt $d.Isp) "isp" (& $each "your isp") $d.Isp.N),
    (& $step "isp international" $d.Intl.Status (& $cnt $d.Intl) "isp-intl" $intlTested $d.Intl.N),
    (& $step $RegionName $d.La.Status (& $cnt $d.La) "region" "all $($d.La.N) mullvad $(Format-Html $run.RefName) servers (listed below)" $d.La.N),
    (& $step $run.CityName $d.Sea.Status (& $cnt $d.Sea) "city" "all $($d.Sea.N) mullvad $(Format-Html $run.CityName) servers (listed below)" $d.Sea.N -Hit:($d.Level -eq "server"))
  )
  foreach ($p in $d.Providers) {
    $servers = Format-Html (Format-Short @($cityRes | Where-Object { $_.Provider -eq $p.Provider }))
    $out += & $step "provider $($p.Provider)" $p.Status (& $cnt $p) "-" $servers $p.N -Hit:($d.Blamed -contains $p.Provider)
  }
  $out -join ""
}

function Get-HeadlineDefinition($run) {
  # What the headline's state means for the level it names; nothing when clean
  $d = $run.Diag
  $sv = $d.Severity
  $rg = { param($n) (Get-Ranges $n).$sv }
  $def = switch ($d.Level) {
    "home" { "$sv means your router lost $(Get-LostPings) or more of $($run.Pings) pings or gave no reply." }
    "isp" { "$sv means $(& $rg $d.Isp.N) of $($d.Isp.N) targets near you are lossy." }
    "isp-intl" { "$sv means $(& $rg $d.Intl.N) of $($d.Intl.N) servers abroad are lossy." }
    "region" { "$sv means $(& $rg $d.Sea.N) of $($d.Sea.N) $($run.CityName) servers are lossy, and $($run.RefName) is at least as bad ($($d.La.Bad) of $($d.La.N))." }
    "city" { "$sv means $(& $rg $d.Sea.N) of $($d.Sea.N) $($run.CityName) servers are lossy, spread across providers." }
    "provider" { (@($d.Providers | Where-Object { $d.Blamed -contains $_.Provider }) | ForEach-Object { "$sv means $(& $rg $_.N) of $($_.N) $(Format-Html $_.Provider) servers are lossy" }) -join "; " }
    "server" { "minor means some $($run.CityName) servers are lossy, but no level above is degraded or a problem." }
    default { $null }
  }
  if ($def) { $def = $def.TrimEnd(".") + ". a server or site is lossy at $BadPct% loss or more ($(Get-LostPings) or more of $($run.Pings) pings) or gives no reply." }
  $def
}

# Favicon: the satellite antenna emoji in an inline svg, percent-encoded so the source stays plain ascii
$Favicon = "data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 viewBox=%270 0 100 100%27%3E%3Ctext y=%27.9em%27 font-size=%2790%27%3E%F0%9F%93%A1%3C/text%3E%3C/svg%3E"

$ReportCss = @'
:root {
  color-scheme: light;
  --bg: #f5f6f8; --fg: #1b1d22; --mut: #646874; --card: #fff; --line: #e4e6eb; --soft: #eef0f4;
  --acc: #4056d6; --accs: #e8ebfc; --onacc: #fff; --shadow: 0 1px 2px rgba(0,0,0,.04); --small: 13px;
  --ok: #23945b; --oks: #e2f4ea; --okt: #1a7446;
  --minor: #7aa83a; --minors: #eef5e0; --minort: #4f7a1c;
  --mixed: #c28a0e; --mixeds: #fbf1d8; --mixedt: #8a5f00;
  --bad: #d2413a; --bads: #fbe3e1; --badt: #b3261e;
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme=light]) {
    color-scheme: dark;
    --bg: #121317; --fg: #e8e9ed; --mut: #9a9eaa; --card: #1b1d23; --line: #2c2f37; --soft: #23262e;
    --acc: #8b9bff; --accs: #252a4a; --onacc: #121317; --shadow: none;
    --ok: #55c98a; --oks: #15301f; --okt: #55c98a;
    --minor: #a6cf6a; --minors: #232e14; --minort: #a6cf6a;
    --mixed: #e5b54c; --mixeds: #352a0f; --mixedt: #e5b54c;
    --bad: #f07a72; --bads: #3b1a18; --badt: #f07a72;
  }
}
.clean { --c: var(--ok); --cs: var(--oks); --ct: var(--okt); }
.minor { --c: var(--minor); --cs: var(--minors); --ct: var(--minort); }
.degraded { --c: var(--mixed); --cs: var(--mixeds); --ct: var(--mixedt); }
.problem { --c: var(--bad); --cs: var(--bads); --ct: var(--badt); }
.n\/a { --c: var(--mut); --cs: var(--soft); --ct: var(--mut); }
* { box-sizing: border-box; }
body { background: var(--bg); color: var(--fg); font: 14px/1.55 system-ui, -apple-system, "Segoe UI", sans-serif; font-variant-numeric: tabular-nums; margin: 0; padding: 28px 16px 48px; }
main { max-width: 760px; margin: auto; }
header { display: flex; align-items: center; justify-content: space-between; gap: 12px 16px; flex-wrap: wrap; }
h2 { font-size: 16px; font-weight: 650; letter-spacing: -.005em; margin: 30px 0 10px; }
.mut { color: var(--mut); }
.card { background: var(--card); border: 1px solid var(--line); border-radius: 14px; padding: 20px 22px; box-shadow: var(--shadow); }
button { font: inherit; font-weight: 600; border: 0; border-radius: 10px; padding: 9px 18px; background: var(--acc); color: var(--onacc); cursor: pointer; }
button:disabled { background: var(--soft); color: var(--mut); cursor: default; }
button.ghost { background: transparent; color: var(--mut); border: 1px solid var(--line); font-weight: 500; }
button.ghost:hover { color: var(--badt); border-color: var(--bad); }
button.small { padding: 5px 12px; font-size: var(--small); }
.runbox { display: flex; align-items: center; gap: 8px 12px; flex-wrap: wrap; margin-left: auto; }
.runbox .mut, .hrow .mut { font-size: var(--small); }
.hero { border-left: 5px solid var(--c); }
.big { font-size: 30px; font-weight: 700; letter-spacing: -.01em; margin: 2px 0 10px; overflow-wrap: anywhere; }
.chips { display: flex; gap: 6px; flex-wrap: wrap; }
.chip { background: var(--soft); border-radius: 999px; padding: 3px 11px; font-size: var(--small); }
.chip.clean, .chip.minor, .chip.degraded, .chip.problem { background: var(--cs); color: var(--ct); }
.sep { margin-top: 16px; padding-top: 14px; border-top: 1px solid var(--line); }
.banner { background: var(--mixeds); color: var(--mixedt); border-radius: 10px; padding: 10px 14px; margin-top: 16px; font-size: var(--small); }
.lead { display: flex; flex-direction: column; align-items: flex-start; gap: 8px; font-size: 15px; }
.pill { display: inline-block; line-height: 1.4; border-radius: 999px; padding: 2px 10px; font-size: var(--small); font-weight: 600; white-space: nowrap; background: var(--cs); color: var(--ct); }
.def { display: block; font-size: var(--small); color: var(--mut); font-style: italic; }
.steps { margin-top: 6px; }
.step { display: grid; grid-template-columns: 22px 1fr auto; gap: 10px; align-items: center; padding: 8px 0; position: relative; }
.step::before { content: ""; position: absolute; left: 10px; top: 0; bottom: 0; width: 2px; background: var(--line); }
.step:first-child::before { top: 50%; }
.step:last-child::before { bottom: 50%; }
.node { width: 12px; height: 12px; border-radius: 50%; margin-left: 5px; position: relative; background: var(--c); box-shadow: 0 0 0 4px var(--card); }
.step.hit .node { width: 16px; height: 16px; margin-left: 3px; box-shadow: 0 0 0 4px var(--card), 0 0 0 7px var(--line); }
.step.hit .lb { font-weight: 700; }
.lb small { display: block; font-size: var(--small); font-weight: 400; color: var(--mut); margin-top: 1px; }
.st { text-align: right; display: flex; flex-direction: column; align-items: flex-end; }
.st b { font-weight: 600; font-size: var(--small); color: var(--ct); }
.st .mut { font-size: var(--small); }
.speed { display: grid; grid-template-columns: minmax(110px, auto) 1fr auto; gap: 10px 14px; align-items: center; }
.speed .val { font-weight: 700; font-size: 16px; text-align: right; }
.speed .sg { grid-column: 1 / -1; font-size: var(--small); font-weight: 600; color: var(--mut); margin-top: 6px; }
.speed .sg:first-child { margin-top: 0; }
.track { height: 10px; border-radius: 5px; background: var(--soft); overflow: hidden; }
.track i { display: block; height: 100%; border-radius: 5px; background: var(--mut); }
.track.acc i { background: var(--acc); }
.note { font-size: var(--small); margin: 14px 0 0; }
.wrap { overflow-x: auto; background: var(--card); border: 1px solid var(--line); border-radius: 14px; box-shadow: var(--shadow); }
table { border-collapse: collapse; width: 100%; }
th, td { padding: 9px 14px; text-align: left; white-space: nowrap; border-bottom: 1px solid var(--line); }
tr:last-child td { border-bottom: 0; }
th { font-weight: 600; color: var(--mut); font-size: var(--small); background: var(--soft); }
.n { text-align: right; }
.nm { font-weight: 600; }
tr.me td { background: var(--accs); }
tr.me td:first-child { box-shadow: inset 4px 0 0 var(--acc); }
td.prob { white-space: normal; min-width: 260px; }
td.prob small { display: block; font-size: var(--small); color: var(--mut); margin-top: 4px; line-height: 1.45; }
.bv { display: inline-flex; align-items: center; gap: 10px; white-space: nowrap; }
.n .bv { justify-content: flex-end; }
.bv .v { min-width: 4.5em; text-align: right; }
.bar { display: inline-block; flex: none; width: 44px; height: 6px; border-radius: 3px; background: var(--soft); overflow: hidden; }
.bar.wide { width: 70px; }
.bar i { display: block; height: 100%; background: var(--c); }
.widesec, .midsec { position: relative; left: 50%; transform: translateX(-50%); width: min(1180px, calc(100vw - 48px)); }
.midsec { width: min(940px, calc(100vw - 48px)); }
.hrow { display: flex; align-items: baseline; gap: 8px 12px; flex-wrap: wrap; margin: 30px 0 10px; }
.hrow h2 { margin: 0; flex: 1; }
@media (max-width: 560px) {
  .step { grid-template-columns: 22px 1fr; }
  .st { grid-column: 2; align-items: flex-start; text-align: left; }
  .widesec, .midsec { left: 0; transform: none; width: auto; }
}
'@

$ReportJs = @'
<script>
(function () {
  var btn = document.getElementById("run");
  var st = document.getElementById("runstate");
  var stop = document.getElementById("stop");
  var fl = document.getElementById("flush");
  var fst = document.getElementById("flushstate");
  var days = fl.getAttribute("data-days");
  var last = null;
  var timer = null;

  function fmt(t) {
    return new Date(t).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }).toLowerCase();
  }

  function show(s) {
    if (s.running) {
      btn.disabled = true;
      btn.textContent = "Running" + String.fromCharCode(8230);
      st.textContent = "started " + fmt(s.started) + ", takes about a minute.";
    } else {
      btn.disabled = false;
      btn.textContent = "Run Now";
      st.textContent = s.next ? "next automatic run " + fmt(s.next) + "." : "";
    }
  }

  function off() {
    btn.disabled = true;
    stop.hidden = true;
    fl.hidden = true;
    st.textContent = "relaycheck-mullvad isn't running. start it with service.cmd.";
  }

  function poll() {
    fetch("/status", { cache: "no-store" })
      .then(function (r) { return r.json(); })
      .then(function (s) {
        if (last !== null && s.last !== last) { location.reload(); return; }
        last = s.last;
        show(s);
      })
      .catch(off);
  }

  function post(path, token) {
    return fetch(path, { method: "POST", headers: { "X-Relaycheck-Mullvad": token } });
  }

  btn.addEventListener("click", function () {
    btn.disabled = true;
    post("/run", "run").then(poll);
  });

  stop.addEventListener("click", function () {
    if (!confirm("stop relaycheck-mullvad? automatic runs stop until you start service.cmd again.")) return;
    post("/stop", "stop").then(function () { clearInterval(timer); off(); });
  });

  fl.addEventListener("click", function () {
    if (!confirm("delete all runs older than " + days + " days? this can't be undone.")) return;
    fl.disabled = true;
    post("/flush", "flush")
      .then(function (r) { return r.json(); })
      .then(function (res) {
        // Drop the flushed rows from this page; the next run rebuilds it from the database anyway
        document.querySelectorAll("tr[data-t]").forEach(function (tr) {
          if (tr.getAttribute("data-t") < res.before) tr.remove();
        });
        fst.textContent = res.removed ? "removed " + res.removed + " run" + (res.removed === 1 ? "" : "s") + " older than " + days + " days." : "nothing older than " + days + " days.";
        fl.disabled = false;
      })
      .catch(function () { fst.textContent = "flush failed."; fl.disabled = false; });
  });

  if (location.protocol === "file:") { off(); return; }
  stop.hidden = false;
  fl.hidden = false;
  poll();
  timer = setInterval(poll, 3000);
})();
</script>
'@

function Write-Report($run) {
  $histSql = "SELECT time, ip, duration_s, severity, problem, diagnosis, pick, direct_down, vpn_down, direct_up, vpn_up, router_vpn FROM runs ORDER BY time DESC LIMIT $HistRows"
  $hist = @(Use-Db { param($db) Get-Rows $db $histSql })
  $d = $run.Diag
  $p = $run.Pick
  $h = $run.Here
  $tc = (Get-Culture).TextInfo

  $off = [TimeZoneInfo]::Local.GetUtcOffset([datetime]::Now)
  $sign = if ($off -lt [TimeSpan]::Zero) { "-" } else { "+" }
  $mins = if ($off.Minutes) { ":{0:00}" -f [math]::Abs($off.Minutes) } else { "" }
  $utc = "utc$sign$([math]::Abs($off.Hours))$mins"
  $when = ([datetime]::ParseExact($run.Time, $TimeFmt, $Inv)).ToString("d MMM yyyy, HH:mm", $Inv).ToLower()
  $place = @($h.City, $h.Country) | Where-Object { $_ }
  $from = @($h.Ip, $h.Isp, ($place -join ", ")) | Where-Object { $_ }
  $from = $from -join " $Dot "
  $headDef = Get-HeadlineDefinition $run

  $sb = New-Object Text.StringBuilder
  [void]$sb.Append("<!doctype html><html lang=en><head><meta charset=utf-8><meta name=viewport content='width=device-width,initial-scale=1'><title>Relaycheck Mullvad</title><link rel=icon href='$Favicon'><style>$ReportCss</style></head><body><main>")
  [void]$sb.Append("<header><div class=mut>$when ($utc) &middot; from $(Format-Html $from)</div>")
  [void]$sb.Append("<div class=runbox><span class=mut id=runstate></span><button id=stop class=ghost hidden>Stop Service</button><button id=run disabled>Run Now</button></div></header>")
  if ($h.ViaMullvad) {
    [void]$sb.Append("<div class=banner>your router&rsquo;s mullvad tunnel is carrying this laptop, so these pings measured the tunnel, not the real route, and this run doesn&rsquo;t count towards the track record. turn it off for this laptop to get a real diagnosis.</div>")
  }

  # Recommended server
  [void]$sb.Append("<h2>Recommended Server</h2>")
  if ($p.Pick) {
    $pp = $p.Pick
    $chips = if ($pp.Ms -eq $null) { "<span class='chip problem'>no reply</span>" }
      else { "<span class='chip $(Get-Tone $pp.Loss $pp.Ms)'>loss $($pp.Loss)%</span><span class=chip>latency $(Format-Ms $pp.Ms)</span>" }
    $heroTone = if ($pp.CleanNow) { "clean" } else { "degraded" }
    [void]$sb.Append("<div class='card hero $heroTone'><div class=big>$(Format-Html $pp.Name)</div>")
    [void]$sb.Append("<div class=chips><span class=chip>provider $(Format-Html $pp.Provider)</span>$chips<span class=chip>clean $($pp.Clean) of $($pp.Runs) runs (${HistHours}h)</span></div>")
    if (-not $pp.CleanNow) { [void]$sb.Append("<p class='note mut'>$(Format-Html $p.Why)</p>") }
    if ($p.Backup) { [void]$sb.Append("<div class=sep><span class=mut>backup</span>&nbsp; <b>$(Format-Html $p.Backup.Name)</b> <span class=mut>&middot; $(Format-Html (Format-Server $p.Backup))</span></div>") }
    [void]$sb.Append("</div>")
  } else {
    [void]$sb.Append("<div class='card hero n/a'><div class=mut>no $(Format-Html $run.CityName) servers could be tested.</div></div>")
  }

  # Where's the problem
  [void]$sb.Append("<h2>Where&rsquo;s The Problem</h2><div class=card><div class=lead><span class='pill $($d.Severity)'>$(Format-Html $d.Label)</span>")
  if ($d.Severity -ne "clean") { [void]$sb.Append("<span>$(Format-Html $d.Text)</span>") }
  if ($headDef) { [void]$sb.Append("<span class=def>$headDef</span>") }
  [void]$sb.Append("</div><div class='sep steps'>$(Format-Ladder $run)</div></div>")

  # Speed
  if ($run.Speed) {
    $s = $run.Speed
    $group = {
      param($title, $dv, $vv)
      $max = [math]::Max(1, [math]::Max([double]$dv, [double]$vv))
      $w = { param($v) if ($v -eq $null) { 0 } else { [math]::Round($v / $max * 100) } }
      "<span class=sg>$title</span><span>direct</span><span class=track><i style='width:$(& $w $dv)%'></i></span><span class=val>$(Format-Mbps $dv)</span>" +
        "<span>vpn via $(Format-Html $p.Pick.Name)</span><span class='track acc'><i style='width:$(& $w $vv)%'></i></span><span class=val>$(Format-Mbps $vv)</span>"
    }
    $lines = @($s.Notes) + @("tested against $($s.Server) on speedtest.net.")
    [void]$sb.Append("<h2>Speed To $($tc.ToTitleCase($run.CityName))</h2><div class=card><div class=speed>")
    [void]$sb.Append((& $group "Download" $s.Direct.Down $s.Vpn.Down) + (& $group "Upload" $s.Direct.Up $s.Vpn.Up) + "</div>")
    [void]$sb.Append("<p class='note mut'>$((@($lines) | ForEach-Object { Format-Html $_ }) -join "<br>")</p></div>")
  }

  # Server tables
  [void]$sb.Append((Format-ServerTable "$($tc.ToTitleCase($run.CityName)) Servers" $p.Ranked $p.Pick.Name))
  [void]$sb.Append((Format-ServerTable "$($tc.ToTitleCase($run.RefName)) Servers" $run.RefRanked $null))

  # Recent runs
  [void]$sb.Append("<section class=widesec><div class=hrow><h2>Recent Runs</h2><span class=mut id=flushstate></span><button id=flush class='ghost small' data-days=$FlushDays hidden>Flush Older Than $FlushDays Days</button></div><div class=wrap><table>")
  [void]$sb.Append("<tr><th>Time</th><th class=n>Run Time</th><th>IP</th><th>Problem</th><th>Recommended</th><th class=n>Direct Download</th><th class=n>VPN Download</th><th class=n>Direct Upload</th><th class=n>VPN Upload</th></tr>")
  foreach ($r in $hist) {
    $t = if ($r.severity) { $r.severity } else { "n/a" }
    $prob = "<span class='pill $t'>$(Format-Html $r.problem)</span>"
    if ($r.diagnosis -and $t -ne "clean") { $prob += "<small>$(Format-Html $r.diagnosis)</small>" }
    if ($r.router_vpn) { $prob += "<small>measured through the router&rsquo;s tunnel</small>" }
    $time = ([datetime]::ParseExact($r.time, $TimeFmt, $Inv)).ToString("d MMM, HH:mm", $Inv)
    $cells = @(
      "<td>$(Format-Html $time)</td>",
      "<td class=n>$(if ($r.duration_s -ne $null) { "$($r.duration_s) s" } else { "&ndash;" })</td>",
      "<td class=mut>$(if ($r.ip) { Format-Html $r.ip } else { "&ndash;" })</td>",
      "<td class=prob>$prob</td>",
      "<td class=nm>$(Format-Html $r.pick)</td>",
      "<td class=n>$(Format-Mbps $r.direct_down)</td>",
      "<td class=n>$(Format-Mbps $r.vpn_down)</td>",
      "<td class=n>$(Format-Mbps $r.direct_up)</td>",
      "<td class=n>$(Format-Mbps $r.vpn_up)</td>"
    )
    [void]$sb.Append("<tr data-t='$(Format-Html $r.time)'>$($cells -join '')</tr>")
  }
  [void]$sb.Append("</table></div></section>$ReportJs</main></body></html>")
  Set-Content $Report $sb.ToString() -Encoding UTF8
}

# ---------------------------------------------------------------- local dashboard (-Serve)
# Serves relaycheck-mullvad.html on http://localhost:<port>; Run Now, Stop Service and Flush post to /run, /stop and /flush.
# Loopback only. Posts need an X-Relaycheck-Mullvad header, which other websites can't send without a cors preflight this
# server never answers, and the Host header must be localhost, which blocks dns rebinding.
function Start-Server([int]$port) {
  $url = "http://localhost:$port"
  $listener = New-Object Net.Sockets.TcpListener ([Net.IPAddress]::Loopback), $port
  try {
    $listener.Start()
  } catch {
    Write-Log "relaycheck-mullvad is already running $Dot opening $url"
    Start-Process $url
    return
  }
  Write-Log "dashboard at $url"
  Start-Process $url

  $st = @{ Child = $null; Started = $null; Next = $null; Stop = $false }
  if ($Loop -gt 0) { $st.Next = Get-Date }

  # Runs started here keep this server's options
  $childArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Pings $Pings -City $City -RefCity $RefCity -RegionName `"$RegionName`""
  if ($NoSpeed) { $childArgs += " -NoSpeed" }

  $startRun = {
    if ($st.Child -and -not $st.Child.HasExited) { return }
    $st.Child = Start-Process powershell -ArgumentList $childArgs -WindowStyle Hidden -PassThru
    $st.Started = Get-Date
    Write-Log "run started (pid $($st.Child.Id))"
  }
  $send = {
    param($stream, $code, $type, [byte[]]$body)
    $hdr = [Text.Encoding]::ASCII.GetBytes("HTTP/1.1 $code`r`nContent-Type: $type`r`nContent-Length: $($body.Length)`r`nCache-Control: no-store`r`nConnection: close`r`n`r`n")
    $stream.Write($hdr, 0, $hdr.Length)
    $stream.Write($body, 0, $body.Length)
  }
  $json = { param($stream, $code, $o) & $send $stream $code "application/json" ([Text.Encoding]::UTF8.GetBytes(($o | ConvertTo-Json -Compress))) }
  $forbid = { param($stream) & $send $stream "403 Forbidden" "text/plain" ([Text.Encoding]::ASCII.GetBytes("forbidden")) }
  $iso = { param($t) if ($t) { $t.ToUniversalTime().ToString("o") } else { $null } }
  $posts = @{ "/run" = "run"; "/stop" = "stop"; "/flush" = "flush" }

  try {
    while (-not $st.Stop) {
      if ($st.Next -and (Get-Date) -ge $st.Next) {
        & $startRun
        $st.Next = (Get-Date).AddMinutes($Loop)
      }
      if (-not $listener.Pending()) {
        Start-Sleep -Milliseconds 150
        continue
      }
      $client = $listener.AcceptTcpClient()
      try {
        # A short read timeout, so a browser preconnect that never sends anything can't stall the loop
        $ns = $client.GetStream()
        $ns.ReadTimeout = 500
        $buf = New-Object byte[] 8192
        $req = ""
        do {
          $n = $ns.Read($buf, 0, $buf.Length)
          $req += [Text.Encoding]::ASCII.GetString($buf, 0, $n)
        } while ($n -gt 0 -and $req -notmatch "`r`n`r`n")

        $lines = $req -split "`r`n"
        $method, $path = ($lines[0] -split " ")[0, 1]
        $hdrs = @{}
        foreach ($l in $lines[1..($lines.Count - 1)]) {
          if ($l -match '^([^:]+):\s*(.*)$') { $hdrs[$matches[1].ToLower()] = $matches[2] }
        }
        if ($hdrs["host"] -notmatch "^(localhost|127\.0\.0\.1):$port$") {
          & $forbid $ns
          continue
        }
        if ($posts.ContainsKey($path) -and ($method -ne "POST" -or $hdrs["x-relaycheck-mullvad"] -ne $posts[$path])) {
          & $forbid $ns
          continue
        }

        switch -regex ($path) {
          '^/(relaycheck-mullvad\.html)?(\?.*)?$' {
            $body = if (Test-Path $Report) { [IO.File]::ReadAllBytes($Report) } else { [Text.Encoding]::UTF8.GetBytes("<p>no report yet. press run now in a moment.</p>") }
            & $send $ns "200 OK" "text/html; charset=utf-8" $body
            break
          }
          '^/status$' {
            $running = [bool]($st.Child -and -not $st.Child.HasExited)
            $last = if (Test-Path $Report) { (Get-Item $Report).LastWriteTimeUtc.ToString("o") } else { $null }
            & $json $ns "200 OK" @{ running = $running; started = (& $iso $st.Started); last = $last; next = (& $iso $st.Next) }
            break
          }
          '^/run$' {
            & $startRun
            & $json $ns "202 Accepted" @{ running = $true }
            break
          }
          '^/stop$' {
            & $json $ns "200 OK" @{ stopped = $true }
            $st.Stop = $true
            Write-Log "stopped from the dashboard"
            break
          }
          '^/flush$' {
            try {
              $f = Remove-OldRuns $FlushDays
              Write-Log "flushed $($f.Removed) runs older than $($f.Before)"
              & $json $ns "200 OK" @{ removed = $f.Removed; before = $f.Before }
            } catch {
              & $json $ns "500 Internal Server Error" @{ error = $_.Exception.Message }
            }
            break
          }
          default {
            & $send $ns "404 Not Found" "text/plain" ([Text.Encoding]::ASCII.GetBytes("not found"))
          }
        }
      } catch {
      } finally {
        $client.Close()
      }
    }
  } finally {
    $listener.Stop()
  }
}

# ---------------------------------------------------------------- console and main loop
function Show-Console($run) {
  $d = $run.Diag
  $p = $run.Pick
  $pad = " " * 13
  $pickLine = if ($p.Pick) { "$($p.Pick.Name) $Dot $($p.Why)" } else { "none" }
  $backupLine = if ($p.Backup) { "$($p.Backup.Name) $Dot $(Format-Server $p.Backup)" } else { "none" }
  $levels = @(@("home network", $d.Home), @("your isp", $d.Isp), @("isp international", $d.Intl), @($RegionName, $d.La), @($run.CityName, $d.Sea)) +
    @($d.Providers | ForEach-Object { , @("provider $($_.Provider)", $_) })

  ""
  "Recommended: $pickLine"
  "Backup:      $backupLine"
  ""
  "Problem:     $($d.Label) $Dot $($d.Text)"
  $pad + (($levels | ForEach-Object { "$($_[0]) $($_[1].Status) ($($_[1].Bad) of $($_[1].N) lossy)" }) -join " $Dot ")
  if ($run.Here.ViaMullvad) { $pad + "note: your router's mullvad tunnel carried these pings, so this run doesn't count." }
  if ($run.Speed) {
    $s = $run.Speed
    $mb = { param($v) if ($v -eq $null) { "no result" } else { "$v mbps" } }
    ""
    "Download:    direct $(& $mb $s.Direct.Down) $Dot vpn via $($p.Pick.Name) $(& $mb $s.Vpn.Down)"
    "Upload:      direct $(& $mb $s.Direct.Up) $Dot vpn via $($p.Pick.Name) $(& $mb $s.Vpn.Up)"
    foreach ($l in $s.Notes) { $pad + $l }
  }
  ""
  "Report:      $Report"
}

if ($Serve) {
  Start-Server $Port
  return
}
do {
  $run = Invoke-Run
  if ($run) {
    Save-Run $run
    Write-Report $run
    Show-Console $run
  }
  if ($Loop -gt 0) {
    Write-Log "next run in $Loop min"
    Start-Sleep -Seconds ($Loop * 60)
  }
} while ($Loop -gt 0)
