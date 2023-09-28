$IdleThreshold = 300000  # 5 Minuten Leerlaufzeit in Millisekunden
$ShutdownDelay = 120  # 2 Minuten Countdown für Warnung in Sekunden

function Is-Idle {
    $LastInputInfo = New-Object -TypeName "IdleTime.NativeMethods+LASTINPUTINFO"
    $LastInputInfo.cbSize = [System.Runtime.InteropServices.Marshal]::SizeOf($LastInputInfo)
    $Result = [IdleTime.NativeMethods]::GetLastInputInfo([ref]$LastInputInfo)
    $IdleTime = ((Get-Date) - [datetime]::FromFileTime($LastInputInfo.dwTime)).TotalMilliseconds
    return $IdleTime -gt $IdleThreshold
}

function Send-Warning {
    [System.Windows.MessageBox]::Show("Ihr PC wird in $($ShutdownDelay / 60) Minuten heruntergefahren. Bitte speichern Sie Ihre Arbeit und fahren Sie den PC manuell herunter, wenn nötig.", "Automatische Abschaltung", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
}

function Shutdown {
    shutdown.exe /s /t $ShutdownDelay
}

Add-Type @"
    using System;
    using System.Runtime.InteropServices;

    public class IdleTime {
        public struct LASTINPUTINFO {
            public uint cbSize;
            public int dwTime;
        }

        public class NativeMethods {
            [DllImport("user32.dll", SetLastError = false)]
            public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
        }
    }
"@ -IgnoreWarnings

while ($true) {
    $CurrentTime = Get-Date
    if ($CurrentTime.DayOfWeek -eq "Friday" -and $CurrentTime.Hour -eq 21) {
        if (Is-Idle) {
            Send-Warning
            Start-Sleep -Seconds $ShutdownDelay
            if (Is-Idle) {
                Shutdown
                break
            }
        }
    }
    Start-Sleep -Seconds 60
}
