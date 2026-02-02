<#
.SYNOPSIS
    Data Review Assistant - Helps navigate through large Excel reports and reference materials
.DESCRIPTION
    Automates routine navigation tasks when reviewing financial reports and data sheets
.NOTES
    Version: 1.2.4
    Company: Internal Tools
#>

Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);

        [DllImport("user32.dll")]
        public static extern bool IsWindow(IntPtr hWnd);
    }
"@

# Minimize console window
$consoleWindow = [Win32]::GetConsoleWindow()
[Win32]::ShowWindow($consoleWindow, 2) | Out-Null

# Global state
$script:excelProcess = $null
$script:chromeProcess = $null
$script:excelWindow = $null
$script:chromeWindow = $null
$script:campaignSheet = $null
$script:isRunning = $true

# Excel COM object
$script:excel = $null

function Write-Status {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Initialize-Applications {
    try {
        # Find Excel
        $script:excelProcess = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $script:excelProcess) {
            Write-Status "Excel not detected. Please open Excel first."
            return $false
        }

        # Connect to Excel COM
        try {
            $script:excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        } catch {
            Write-Status "Unable to connect to Excel."
            return $false
        }

        # Find Campaign sheet
        $found = $false
        foreach ($wb in $script:excel.Workbooks) {
            foreach ($ws in $wb.Worksheets) {
                if ($ws.Name -like "*Campaign*") {
                    $script:campaignSheet = $ws
                    $ws.Activate()
                    $found = $true
                    break
                }
            }
            if ($found) { break }
        }

        $script:excelWindow = $script:excelProcess.MainWindowHandle

        # Find Chrome (optional)
        $script:chromeProcess = Get-Process -Name "chrome" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($script:chromeProcess) {
            $script:chromeWindow = $script:chromeProcess.MainWindowHandle
        }

        return $true
    } catch {
        Write-Status "Initialization issue. Retrying..."
        return $false
    }
}

function Invoke-SafeActivate {
    param([IntPtr]$WindowHandle)

    try {
        if ([Win32]::IsWindow($WindowHandle)) {
            [Win32]::SetForegroundWindow($WindowHandle) | Out-Null
            Start-Sleep -Milliseconds 50
            return $true
        }
    } catch {
        # Silent fail
    }
    return $false
}

function Send-SafeKeys {
    param([string]$Keys)

    try {
        $wshell = New-Object -ComObject WScript.Shell
        $wshell.SendKeys($Keys)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wshell) | Out-Null
    } catch {
        # Silent fail
    }
}

function Clear-Dialogs {
    Send-SafeKeys("{ESC}")
    Start-Sleep -Milliseconds 100
    Send-SafeKeys("{ESC}")
    Start-Sleep -Milliseconds 100
    Send-SafeKeys("{ESC}")
    Start-Sleep -Milliseconds 150
}

function Invoke-ExcelNavigation {
    if (-not (Invoke-SafeActivate $script:excelWindow)) {
        return
    }

    Clear-Dialogs

    $actions = @(
        # Arrow navigation
        { Send-SafeKeys("{DOWN}") },
        { Send-SafeKeys("{DOWN}") },
        { Send-SafeKeys("{RIGHT}") },
        { Send-SafeKeys("{UP}") },
        { Send-SafeKeys("{LEFT}") },

        # Page navigation
        { Send-SafeKeys("{PGDN}") },
        { Send-SafeKeys("{PGUP}") },

        # Cell jumps
        { Send-SafeKeys("^{HOME}") },
        { Send-SafeKeys("^{DOWN}") },
        { Send-SafeKeys("^{RIGHT}") },

        # Ctrl+G navigation with row numbers
        {
            try {
                $row = Get-Random -Minimum 100 -Maximum 32000000
                Send-SafeKeys("^g")
                Start-Sleep -Milliseconds 200
                Send-SafeKeys("A$row")
                Start-Sleep -Milliseconds 100
                Send-SafeKeys("{ENTER}")
            } catch {}
        },

        # Column scrolling
        {
            $col = [char](Get-Random -Minimum 65 -Maximum 117)
            Send-SafeKeys("^g")
            Start-Sleep -Milliseconds 200
            Send-SafeKeys("$col" + "1")
            Start-Sleep -Milliseconds 100
            Send-SafeKeys("{ENTER}")
        },

        # Multi-direction navigation
        {
            1..3 | ForEach-Object {
                Send-SafeKeys("{DOWN}")
                Start-Sleep -Milliseconds 50
            }
        },
        {
            1..5 | ForEach-Object {
                Send-SafeKeys("{RIGHT}")
                Start-Sleep -Milliseconds 50
            }
        },

        # End key navigation
        { Send-SafeKeys("^{END}") },
        { Send-SafeKeys("{HOME}") }
    )

    # Execute 5 random actions
    1..5 | ForEach-Object {
        try {
            $action = $actions | Get-Random
            & $action
            Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 600)
        } catch {
            # Continue on error
        }
    }

    Clear-Dialogs
}

function Invoke-ChromeView {
    if (-not $script:chromeWindow -or -not (Invoke-SafeActivate $script:chromeWindow)) {
        return
    }

    Clear-Dialogs

    $actions = @(
        # Scroll down
        { Send-SafeKeys("{PGDN}") },

        # Scroll up
        { Send-SafeKeys("{PGUP}") },

        # Arrow scroll
        { Send-SafeKeys("{DOWN}") },
        { Send-SafeKeys("{UP}") },

        # Smooth scroll
        {
            1..3 | ForEach-Object {
                Send-SafeKeys("{DOWN}")
                Start-Sleep -Milliseconds 100
            }
        },

        # Space bar scroll
        { Send-SafeKeys(" ") }
    )

    # Execute 3-5 actions
    $count = Get-Random -Minimum 3 -Maximum 6
    1..$count | ForEach-Object {
        try {
            $action = $actions | Get-Random
            & $action
            Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 600)
        } catch {
            # Continue on error
        }
    }

    Clear-Dialogs
}

function Start-ReviewCycle {
    while ($script:isRunning) {
        try {
            # Verify processes still exist
            if (-not (Get-Process -Id $script:excelProcess.Id -ErrorAction SilentlyContinue)) {
                Write-Status "Excel closed."
                break
            }

            # Decide: Excel (80%) or Chrome (20%)
            $useExcel = (Get-Random -Minimum 1 -Maximum 101) -le 80

            if ($useExcel) {
                Invoke-ExcelNavigation
            } else {
                if ($script:chromeWindow) {
                    Invoke-ChromeView
                } else {
                    Invoke-ExcelNavigation
                }
            }

            # Pause between bursts
            $pauseSeconds = Get-Random -Minimum 1 -Maximum 16
            Start-Sleep -Seconds $pauseSeconds

        } catch {
            Start-Sleep -Seconds 2
            continue
        }
    }
}

# Cleanup handler
function Stop-ReviewCycle {
    $script:isRunning = $false

    if ($script:excel) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:excel) | Out-Null
        } catch {}
    }

    Write-Status "Session complete."
}

# Trap Ctrl+C
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    Stop-ReviewCycle
}

try {
    # Header
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "  Data Review Assistant v1.2.4" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""

    Write-Status "Initializing..."

    if (-not (Initialize-Applications)) {
        Write-Host ""
        Write-Host "Press any key to exit..."
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }

    Write-Status "Connected to Excel."
    if ($script:campaignSheet) {
        Write-Status "Campaign sheet activated."
    }
    if ($script:chromeWindow) {
        Write-Status "Chrome reference detected."
    }

    Write-Host ""
    Write-Status "Ready. Press Ctrl+C to stop."
    Write-Host ""

    Start-Sleep -Seconds 2

    # Main loop
    Start-ReviewCycle

} catch {
    Write-Status "An error occurred. Exiting."
} finally {
    Stop-ReviewCycle
}
