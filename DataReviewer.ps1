# Data Review Assistant v1.2.4
# Excel report navigation helper

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, IntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_WHEEL = 0x0800;
}
"@

[Win32]::ShowWindow([Win32]::GetConsoleWindow(), 0)

try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application") } 
catch { exit }

if ($excel.Workbooks.Count -eq 0) { exit }
$wb = $excel.ActiveWorkbook
$ws = $wb.ActiveSheet
$wshell = New-Object -ComObject wscript.shell

# Perpetua pages to navigate
$perpetuaPages = @(
    "https://app.perpetua.io/am/prism/niches/manage/B0D57PMKZJ?company=45134&geocompany=58088",
    "https://app.perpetua.io/am/prism/market-segment?company=45134&geocompany=58088&range=THIRTY_DAYS",
    "https://app.perpetua.io/am/sp/?company=45134&geocompany=58088",
    "https://app.perpetua.io/am/brands-v4/?range=THIRTY_DAYS&company=45134&geocompany=58088",
    "https://app.perpetua.io/am/sd?company=45134&geocompany=58088&tabId=goals"
)

$currentPageIndex = 0

function Scroll-It { 
    param([int]$d=-3)
    for ($i=0; $i -lt [Math]::Abs($d); $i++) {
        [Win32]::mouse_event(0x0800,0,0,$(if($d -lt 0){-120}else{120}),[IntPtr]::Zero)
        Start-Sleep -Milliseconds (Get-Random -Min 80 -Max 200)
    }
}

function Close-Dialogs {
    1..3 | %{ [System.Windows.Forms.SendKeys]::SendWait("{ESC}"); Start-Sleep -Milliseconds 50 }
}

function Navigate-PerpetuaPage {
    # Navigate to next Perpetua page
    [System.Windows.Forms.SendKeys]::SendWait("^l")
    Start-Sleep -Milliseconds 300
    [System.Windows.Forms.SendKeys]::SendWait("^a")
    Start-Sleep -Milliseconds 200
    
    $page = $perpetuaPages[$script:currentPageIndex]
    foreach ($char in $page.ToCharArray()) {
        [System.Windows.Forms.SendKeys]::SendWait($char)
        Start-Sleep -Milliseconds (Get-Random -Min 20 -Max 80)
    }
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    
    $script:currentPageIndex = ($script:currentPageIndex + 1) % $perpetuaPages.Count
    Start-Sleep -Milliseconds (Get-Random -Min 2000 -Max 4000)
}

$cols = @("A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ")
$rows = @(100,500,1000,5000,10000,50000,100000,500000,1000000,5000000,10000000,20000000,30000000)
$apps = @("excel","excel","excel","excel","chrome")

$chromePageCounter = 0

while ($true) {
    $app = $apps | Get-Random
    
    if ($app -eq "excel") {
        $wshell.AppActivate("Excel") | Out-Null
    } else {
        $wshell.AppActivate("Chrome") | Out-Null
    }
    
    Start-Sleep -Milliseconds 300
    Close-Dialogs
    
    for ($i=0; $i -lt 5; $i++) {
        if ($app -eq "excel") {
            try {
                switch (Get-Random -Min 1 -Max 9) {
                    1 { $ws.Range("$($cols|Get-Random)$($rows|Get-Random)").Select()|Out-Null }
                    2 { Scroll-It (Get-Random -Min -8 -Max -2) }
                    3 { Scroll-It (Get-Random -Min 2 -Max 8) }
                    4 { [System.Windows.Forms.SendKeys]::SendWait("{PGDN}") }
                    5 { [System.Windows.Forms.SendKeys]::SendWait("{PGUP}") }
                    6 { [System.Windows.Forms.SendKeys]::SendWait("$('{RIGHT}'*$(Get-Random -Min 2 -Max 5))") }
                    7 { [System.Windows.Forms.SendKeys]::SendWait("$('{LEFT}'*$(Get-Random -Min 2 -Max 5))") }
                    8 { [System.Windows.Forms.SendKeys]::SendWait("^{HOME}"); Start-Sleep -Ms 200; $ws.Range("A$($rows|Get-Random)").Select()|Out-Null }
                }
            } catch {}
        } else {
            # Chrome: Navigate Perpetua pages + scroll
            if ($chromePageCounter % 10 -eq 0) {
                Navigate-PerpetuaPage
            }
            
            $x = Get-Random -Min 200 -Max 1200
            $y = Get-Random -Min 200 -Max 700
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
            Scroll-It (Get-Random -Min -5 -Max 5)
            $chromePageCounter++
        }
        
        Start-Sleep -Milliseconds (Get-Random -Min 200 -Max 600)
    }
    
    Close-Dialogs
    Start-Sleep -Seconds (Get-Random -Min 1 -Max 15)
}
