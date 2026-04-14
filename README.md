# EVE_DPS-counter
A simple Windows tool for reading EVE Online combat logs and showing:
- Current Incoming / Outgoing DPS
- Total  Incoming / Outgoing DPS
- Total  Incoming / Outgoing damage
- Recent combat graph

## Features
- Auto-detects running EVE characters
- Tracks outgoing damage from combat logs
- Uses incoming or outgoing damage to keep combat status active
- Excludes non-attacking time from DPS duration
- Supports character alias and hide
- Shows recent combat graphs and battle history
- Supports taskbar minimize
- Incoming damage tracking (Inc.DPS, Inc.T.DPS, Inc.T.Dam)
- Customizable character sorting — Top/Bottom dealer, A→Z, or drag-and-drop manual order
- DPS bar graph overlay on each character row
- Resizable graph/history window with persistent size and position
- Alarm system — visual blink and sound alert when Inc.DPS exceeds a set threshold
- Supports English and Korean EVE clients (JP, DE, RU, FR also supported)

## Requirements
- Windows
- EVE Online

## Download
Download the latest **EVE_DPS.exe** from the release page(https://github.com/opteryx7/EVE_DPS-counter/releases). No installation or Python required.

## Security Notice
Some antivirus tools may flag this as suspicious due to the PyInstaller packaging method — this is a known false positive. If blocked, add an exception in your antivirus or Windows Defender settings.

This tool is fully open source — you can review every line of code on GitHub. It only reads EVE Online's local combat log files and saves a small config file to your home directory. No network connections, no data collection, no external servers.

[VirusTotal scan results] (https://www.virustotal.com/gui/file/5a51eca36e753ca3b2bf3bd36e426737958f75f4c290993bf7668c49bc6b84cd?nocache=1)

## Changelog
**v1.3**
- Incoming damage panel (Inc.DPS / Inc.T.DPS / Inc.T.Dam) with collapse toggle
- Character sort order (Top/Bottom dealer, name, manual drag-and-drop) — persists on restart
- DPS bar graph background per character row
- Resizable and repositionable graph/history window
- Alarm system with visual blink and sound alert (customizable threshold and WAV)
- Graph window title bar removed for compact layout
- History scrollbar styled to match UI
- Multi-language combat log support (EN/KO/JP/DE/RU/FR)

**v1.2**
- Battle history tab with per-character DPS and damage
- Recent battle graph
- Character alias (inline edit) and hide
- Window position and transparency saved on exit
- Korean client support

## Run from source
```cmd
pip install pywin32
py EVE_DPS.py
```
