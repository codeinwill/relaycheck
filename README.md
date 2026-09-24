# Relaycheck 📡

Relaycheck tells you which Mullvad VPN server to use, and when the connection is bad, where the problem is: your home network, your ISP, your ISP's international links, the route to the region, the city, one hosting provider, or a few individual servers.

It runs on Windows, detects your network on every run (location, ISP, home router, your ISP's first router), and shows the results on a local dashboard with a Run Now button.

By default it checks Mullvad's Seattle servers and uses Los Angeles as the West Coast reference. Any Mullvad city can be tested.

## Quick start

1. Double-click `service.cmd`.
2. Your browser opens **http://localhost:8765**. The first check runs straight away and takes about a minute.
3. After that it checks again every 30 minutes.

`service.cmd` starts Relaycheck in the background with no window. To stop it, press **Stop Service** on the dashboard. Closing the browser tab doesn't stop it. Double-clicking `service.cmd` while it's running just opens the dashboard again.

## Requirements

- Windows 10 or 11, with Windows PowerShell 5.1 (built in)
- The [Mullvad VPN app](https://mullvad.net/download) for the VPN speed test. Without it, Relaycheck still runs the ping checks and the direct speed test.
- Nothing to install: the database uses the SQLite library that ships with Windows (`winsqlite3.dll`).

## What each run does

1. **Pings a ladder of targets**, 50 pings each, all at once:

   | Level | What's tested |
   |---|---|
   | home network | your router (the default gateway) |
   | your isp | your ISP's first router, 1.1.1.1 and 8.8.8.8 (both answer from near you) |
   | isp international | the 2 nearest Mullvad servers outside your country |
   | us west coast | every Mullvad server in the reference city (Los Angeles) |
   | seattle | every Mullvad server in the city under test |
   | provider | the city's servers grouped by hosting company |

2. **Gives each level a state** from how many of its targets are lossy. A target is lossy at 4% loss or more (2 or more of 50 pings lost), or when it gives no reply.

   | State | Share of the level's targets that are lossy |
   |---|---|
   | clean | none |
   | minor | up to 25% |
   | degraded | more than 25%, less than 60% |
   | problem | 60% or more |

   The headline is the first level from the top that is a problem, else the first that is degraded, else any lossy servers on their own (minor). A provider is only named when 2 or more of its servers are lossy and another provider is healthier. A router that never answers pings is shown as "ignores ping" and isn't counted.

3. **Recommends a server:** one that's clean now and has the best track record over the last 24 hours, plus a backup from a different provider.

4. **Runs a speed test** on speedtest.net, the same server each time: first direct, then through the recommended server. It measures download and upload.

5. **Saves the results** to `relaycheck.db` and rebuilds `relaycheck.html`.

The laptop's Mullvad app is disconnected while pinging, so the pings take the real route. It's connected to the recommended server only for the speed test, then put back the way it was. If your router runs a Mullvad tunnel that carries the laptop's traffic, the page warns you, and that run doesn't count towards the track record.

## Dashboard

- **Recommended Server:** the server to use and a backup, with loss, latency and 24h track record.
- **Where's The Problem:** the headline and the full ladder. Each level shows what was tested and what its state means.
- **Speed To Seattle:** download and upload, direct and through the VPN.
- **Seattle Servers and Los Angeles Servers:** every server's loss, latency, jitter, and how often it was clean over 24 hours and 7 days.
- **Recent Runs:** history, and a **Flush Older Than 7 Days** button.

## Running it by hand

```
powershell -ExecutionPolicy Bypass -File relaycheck.ps1                          one run
powershell -ExecutionPolicy Bypass -File relaycheck.ps1 -NoSpeed                 skip the speed test
powershell -ExecutionPolicy Bypass -File relaycheck.ps1 -Loop 30                 run, wait 30 min, repeat
powershell -ExecutionPolicy Bypass -File relaycheck.ps1 -Serve -Loop 30          dashboard, run every 30 min (what service.cmd does)
powershell -ExecutionPolicy Bypass -File relaycheck.ps1 -City lax -RefCity sea   test another Mullvad city
```

| Option | Default | Meaning |
|---|---|---|
| `-City` | `sea` | Mullvad city code to test |
| `-RefCity` | `lax` | Mullvad city code used as the regional reference |
| `-RegionName` | `us west coast` | label for the reference level |
| `-Pings` | `50` | pings per target |
| `-Loop` | `0` | minutes between runs (0 = run once) |
| `-Serve` | off | start the dashboard on localhost |
| `-Port` | `8765` | dashboard port |
| `-NoSpeed` | off | skip the speed test |

The thresholds (clean and lossy loss, state shares, history windows) are at the top of `relaycheck.ps1`.

## Files

| File | What it is |
|---|---|
| `service.cmd` | starts the dashboard in the background |
| `relaycheck.ps1` | the whole program |
| `relaycheck.db` | results (SQLite), created on the first run |
| `relaycheck.html` | the latest report, rebuilt on every run |

`relaycheck.db` and `relaycheck.html` hold your public IP, ISP and location, so they're excluded from git.

### Database

Open `relaycheck.db` with [DB Browser for SQLite](https://sqlitebrowser.org).

- `runs`: one row per run. Where you were, the diagnosis, the recommended server and backup, and speed results.
- `checks`: one row per target per run. Level (`grp`), name, provider, IP, loss, latency and jitter.

## Dashboard security

The dashboard only listens on `127.0.0.1`. Run Now, Stop Service and Flush need a custom request header that other websites can't send without a CORS preflight, which the server never answers. Requests whose `Host` header isn't localhost are refused, which blocks DNS rebinding.
