# EVE_DPS-counter

A simple Windows tool for reading EVE Online combat logs and showing:

- Current DPS
- Total DPS
- Total damage
- Recent combat graph

## Features

- Auto-detects running EVE characters
- Tracks outgoing damage from combat logs
- Uses incoming or outgoing damage to keep combat status active
- Excludes non-attacking time from DPS duration
- Supports character alias and hide
- Shows recent combat graphs for visible characters
- Supports taskbar minimize

## Requirements

- Windows
- Python 3
- EVE Online combat logs
- `pywin32`

## Run

```cmd
py EVE_DPS.py
