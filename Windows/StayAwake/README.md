# StayAwake PowerShell Script

A PowerShell script that prevents your Windows system from going to sleep by simulating user activity at specified intervals.

## Prerequisites

- Windows operating system
- PowerShell 5.1 or later
- Administrator privileges (for some features)

## Quick Start

1. Download `StayAwake.ps1`
2. Open PowerShell as administrator
3. Navigate to script directory
4. Run: `.\StayAwake.ps1`

## Command-line Options

- `-Interval`: Time in seconds between key simulations (Default: 240)
- `-Duration`: How long to run in minutes (Default: runs indefinitely)
- `-Key`: Key to simulate (Default: Scroll Lock)
- `-Quiet`: Suppress console output

Example:
```powershell
.\StayAwake.ps1 -Interval 300 -Duration 60 -Key "NumLock" -Quiet
```

## Installation

1. Download the script
2. Place in desired location
3. Optionally add to PowerShell profile for easy access

## How It Works

The script:
1. Simulates keypress at specified intervals
2. Uses Windows API for key simulation
3. Runs until manually stopped or duration reached
4. Maintains system awake state

## Customization

Modify script variables to:
- Change default interval
- Set different keys
- Adjust output formatting
- Add custom actions

## Troubleshooting

Common issues:
- Script not running: Check execution policy
- Admin rights needed: Run as administrator
- Key not working: Try different key option

## Contributing

1. Fork repository
2. Create feature branch
3. Submit pull request
4. Follow coding standards

## License

MIT License - Feel free to use and modify