param (
    [string]$process,
    [string]$exe
)

Add-Type @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class WinAPI {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    
    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
    
    [DllImport("kernel32.dll")]
    public static extern int GetCurrentThreadId();
    
    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
    
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    
    // Virtual Desktop Detection
    [DllImport("dwmapi.dll")]
    public static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out bool pvAttribute, int cbAttribute);
    
    public const int DWMWA_CLOAKED = 14;
}
"@

function Test-WindowOnCurrentDesktop {
    param([IntPtr]$hwnd)
    
    if ($hwnd -eq [IntPtr]::Zero) { return $false }
    if (-not [WinAPI]::IsWindowVisible($hwnd)) { return $false }
    
    # Prüfe ob das Fenster "cloaked" ist (versteckt auf anderem Desktop)
    try {
        $cloaked = $false
        $result = [WinAPI]::DwmGetWindowAttribute($hwnd, [WinAPI]::DWMWA_CLOAKED, [ref]$cloaked, 4)
        if ($result -eq 0) {
            if ($cloaked) {
                Write-Host "Fenster ist cloaked (anderer Desktop)" -ForegroundColor Red
                return $false  # Fenster ist definitiv auf anderem Desktop
            } else {
                Write-Host "Fenster ist nicht cloaked (aktueller Desktop)" -ForegroundColor Green
                return $true   # Fenster ist definitiv auf aktuellem Desktop
            }
        }
    } catch {
        Write-Host "DWM-Check fehlgeschlagen" -ForegroundColor Yellow
    }
    
    # Fallback: Teste durch versuchtes Aktivieren
    $currentForeground = [WinAPI]::GetForegroundWindow()
    $testResult = [WinAPI]::SetForegroundWindow($hwnd)
    Start-Sleep -Milliseconds 50
    $newForeground = [WinAPI]::GetForegroundWindow()
    
    # Stelle ursprünglichen Zustand wieder her, falls nötig
    if ($currentForeground -ne [IntPtr]::Zero -and $newForeground -ne $hwnd) {
        [WinAPI]::SetForegroundWindow($currentForeground)
    }
    
    $isOnCurrentDesktop = ($newForeground -eq $hwnd)
    Write-Host "Fallback-Test: Fenster $(if($isOnCurrentDesktop){'ist'}else{'ist nicht'}) auf aktuellem Desktop" -ForegroundColor $(if($isOnCurrentDesktop){'Green'}else{'Red'})
    
    return $isOnCurrentDesktop
}

function Get-ProcessesOnCurrentDesktop {
    param([string[]]$processNames)
    
    $currentDesktopProcesses = @()
    
    foreach ($name in $processNames) {
        $processes = Get-Process -Name $name -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
        
        Write-Host "Prozess '$name': $($processes.Count) Instanzen gefunden" -ForegroundColor Gray
        
        foreach ($proc in $processes) {
            Write-Host "  Prüfe PID $($proc.Id), Handle: $($proc.MainWindowHandle)" -ForegroundColor Gray
            
            # Spezielle Behandlung für OneNote - prüfe auch Fenster ohne MainWindowTitle
            $isValidWindow = $true
            if ($name -match "ONENOTE") {
                # Für OneNote: Auch Fenster ohne Titel akzeptieren
                Write-Host "  OneNote erkannt - erweiterte Fensterprüfung" -ForegroundColor Yellow
            } else {
                # Für andere Apps: Fenster müssen einen Titel haben
                if ([string]::IsNullOrEmpty($proc.MainWindowTitle)) {
                    Write-Host "  Überspringe: Kein Fenstertitel" -ForegroundColor Gray
                    $isValidWindow = $false
                }
            }
            
            if ($isValidWindow -and (Test-WindowOnCurrentDesktop -hwnd $proc.MainWindowHandle)) {
                Write-Host "  → Auf aktuellem Desktop!" -ForegroundColor Green
                $currentDesktopProcesses += $proc
            } else {
                Write-Host "  → Nicht auf aktuellem Desktop" -ForegroundColor Red
            }
        }
    }
    
    return $currentDesktopProcesses
}

function Get-AllProcesses {
    param([string[]]$processNames)
    
    $allProcesses = @()
    foreach ($name in $processNames) {
        $processes = Get-Process -Name $name -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }
        Write-Host "Alle Prozesse '$name': $($processes.Count) gefunden" -ForegroundColor Gray
        $allProcesses += $processes
    }
    return $allProcesses
}

# Mehrere mögliche Namen erlauben (durch Komma getrennt)
$procNames = $process -split "," | ForEach-Object { $_.Trim() }

Write-Host "Suche nach Prozessen: $($procNames -join ', ')" -ForegroundColor Cyan

# Suche zuerst nach Prozessen auf dem aktuellen Desktop
$currentDesktopProcs = Get-ProcessesOnCurrentDesktop -processNames $procNames

Write-Host "Gefundene Prozesse auf aktuellem Desktop: $($currentDesktopProcs.Count)" -ForegroundColor Cyan

if ($currentDesktopProcs.Count -gt 0) {
    # Es gibt ein Fenster auf dem aktuellen Desktop
    $proc = $currentDesktopProcs[0]
    $hwnd = $proc.MainWindowHandle
    $activeHwnd = [WinAPI]::GetForegroundWindow()
    
    if ($hwnd -eq $activeHwnd) {
        # Fenster ist bereits aktiv → minimieren
        [WinAPI]::ShowWindowAsync($hwnd, 6) # SW_MINIMIZE
        Write-Host "Fenster minimiert" -ForegroundColor Green
        exit
    }
    
    if ([WinAPI]::IsIconic($hwnd)) {
        # Fenster ist minimiert → wiederherstellen
        [WinAPI]::ShowWindowAsync($hwnd, 9) # SW_RESTORE
    }
    
    # Fenster aktivieren (sicher, da es auf dem aktuellen Desktop ist)
    [WinAPI]::SetForegroundWindow($hwnd)
    Write-Host "Fenster aktiviert" -ForegroundColor Green
    exit
}

# Prüfe ob Prozesse auf anderen Desktops laufen
$allProcesses = Get-AllProcesses -processNames $procNames

Write-Host "Alle gefundenen Prozesse: $($allProcesses.Count)" -ForegroundColor Cyan

if ($allProcesses.Count -gt 0) {
    Write-Host "App läuft auf anderem Desktop. Starte neue Instanz im aktuellen Desktop..." -ForegroundColor Yellow
} else {
    Write-Host "App läuft nicht. Starte neue Instanz..." -ForegroundColor Cyan
}

# Neue Instanz im aktuellen Desktop starten
try {
    # Spezielle Behandlung für OneNote
    if ($exe -match "ONENOTE") {
        Write-Host "OneNote-spezifischer Start (Single-Instance Problem)..." -ForegroundColor Yellow
        
        # Für OneNote: Versuche verschiedene Methoden
        $oneNoteStarted = $false
        
        # Methode 1: Mit explorer.exe starten (Windows Shell Kontext)
        try {
            Write-Host "Versuche Start über explorer.exe..." -ForegroundColor Cyan
            Start-Process -FilePath "explorer.exe" -ArgumentList "`"$exe`"" -WindowStyle Hidden
            Start-Sleep -Milliseconds 500
            $oneNoteStarted = $true
        } catch {
            Write-Host "Explorer-Start fehlgeschlagen" -ForegroundColor Red
        }
        
        # Methode 2: Mit PowerShell Job (isolierter Kontext)
        if (-not $oneNoteStarted) {
            try {
                Write-Host "Versuche Start über PowerShell Job..." -ForegroundColor Cyan
                $job = Start-Job -ScriptBlock {
                    param($exePath)
                    Start-Process -FilePath $exePath -WindowStyle Normal
                } -ArgumentList $exe
                
                Wait-Job $job -Timeout 5 | Remove-Job
                $oneNoteStarted = $true
            } catch {
                Write-Host "PowerShell Job fehlgeschlagen" -ForegroundColor Red
            }
        }
        
        # Methode 3: Mit WScript Shell (COM)
        if (-not $oneNoteStarted) {
            try {
                Write-Host "Versuche Start über WScript.Shell COM..." -ForegroundColor Cyan
                $shell = New-Object -ComObject WScript.Shell
                $shell.Run("`"$exe`"", 1, $false)
                $oneNoteStarted = $true
            } catch {
                Write-Host "WScript.Shell fehlgeschlagen" -ForegroundColor Red
            }
        }
        
        if ($oneNoteStarted) {
            Write-Host "OneNote Start-Versuch abgeschlossen" -ForegroundColor Green
        } else {
            Write-Host "Alle OneNote-Startmethoden fehlgeschlagen - Fallback" -ForegroundColor Red
            Start-Process -FilePath $exe
        }
        
    } else {
        # Für andere Apps: cmd.exe als Launcher verwenden
        $cmdArgs = "/c `"$exe`""
        
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = "cmd.exe"
        $startInfo.Arguments = $cmdArgs
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
        
        $cmdProcess = [System.Diagnostics.Process]::Start($startInfo)
        
        if ($cmdProcess) {
            Write-Host "Neue Instanz gestartet im aktuellen Desktop" -ForegroundColor Green
            $cmdProcess.WaitForExit()
        }
    }
} catch {
    if ($exe -match "ONENOTE") {
       Write-Host "Starte OneNote direkt via Explorer AppUserModelID auf aktuellem Desktop..."
       Start-Process "explorer.exe" -ArgumentList "shell:AppsFolder\Microsoft.Office.OneNote_8wekyb3d8bbwe!OneNote"
    }
}
