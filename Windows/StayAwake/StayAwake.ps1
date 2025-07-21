Add-Type @"
using System;
using System.Runtime.InteropServices;
public class PowerState {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
"@

# Safely parse the flags as UInt32
$ES_CONTINUOUS = [System.UInt32]::Parse("2147483648")
$ES_SYSTEM_REQUIRED = [System.UInt32]::Parse("1")
$ES_DISPLAY_REQUIRED = [System.UInt32]::Parse("2")

# Combine flags
$esFlags = $ES_CONTINUOUS -bor $ES_SYSTEM_REQUIRED -bor $ES_DISPLAY_REQUIRED

while ($true) {
    [PowerState]::SetThreadExecutionState($esFlags) | Out-Null
    Start-Sleep -Seconds 300  # 5 minutes
}
