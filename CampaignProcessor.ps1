# Campaign Data Processor v1.0
# Excel workflow helper

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Mouse {
    [DllImport("user32.dll")]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, IntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_LEFTDOWN = 0x02;
    public const uint MOUSEEVENTF_LEFTUP = 0x04;
    public const uint MOUSEEVENTF_WHEEL = 0x0800;
}
public class Console {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

# Hide this PowerShell window
$consolePtr = [Console]::GetConsoleWindow()
[Console]::ShowWindow($consolePtr, 2)  # 2 = Minimize

# ============== CONFIGURATION ==============

$script:screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$script:screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height

# Excel COM object (for real file operations)
$script:excel = $null
$script:workbook = $null
$script:worksheet = $null
$script:bulkFileLoaded = $false

# Task history for anti-pattern (never repeat last 20)
$script:taskHistory = @()
$script:lastApp = ""
$script:sessionStart = Get-Date

# Bulk file metadata (populated after loading)
$script:bulkFileInfo = @{
    totalRows = 0
    worksheets = @()
    columns = @()
    rowRanges = @()
}

# ============== 140+ MICRO-TASKS ==============

$script:microTasks = @(
    # === EXCEL - CAMPAIGN DATA (25 tasks) ===
    @{ id=1;  cat="excel"; name="Enter SKU in cell"; dur=@(8,25); weight=4 }
    @{ id=2;  cat="excel"; name="Type campaign name"; dur=@(12,35); weight=4 }
    @{ id=3;  cat="excel"; name="Enter ACOS value"; dur=@(5,15); weight=5 }
    @{ id=4;  cat="excel"; name="Type spend amount"; dur=@(5,12); weight=5 }
    @{ id=5;  cat="excel"; name="Enter sales figure"; dur=@(5,12); weight=4 }
    @{ id=6;  cat="excel"; name="Type impressions count"; dur=@(6,15); weight=3 }
    @{ id=7;  cat="excel"; name="Enter clicks value"; dur=@(5,12); weight=3 }
    @{ id=8;  cat="excel"; name="Type CTR percentage"; dur=@(5,12); weight=3 }
    @{ id=9;  cat="excel"; name="Enter CPC value"; dur=@(5,12); weight=3 }
    @{ id=10; cat="excel"; name="Type conversion rate"; dur=@(6,15); weight=3 }
    @{ id=11; cat="excel"; name="Add ROAS formula"; dur=@(15,40); weight=3 }
    @{ id=12; cat="excel"; name="Create SUM formula"; dur=@(10,30); weight=3 }
    @{ id=13; cat="excel"; name="Add AVERAGE formula"; dur=@(10,30); weight=2 }
    @{ id=14; cat="excel"; name="Enter date value"; dur=@(8,20); weight=3 }
    @{ id=15; cat="excel"; name="Type ad group name"; dur=@(10,30); weight=3 }
    @{ id=16; cat="excel"; name="Enter keyword text"; dur=@(12,35); weight=4 }
    @{ id=17; cat="excel"; name="Type match type"; dur=@(5,12); weight=3 }
    @{ id=18; cat="excel"; name="Add targeting note"; dur=@(15,45); weight=2 }
    @{ id=19; cat="excel"; name="Enter budget value"; dur=@(5,15); weight=3 }
    @{ id=20; cat="excel"; name="Type bid amount"; dur=@(5,12); weight=4 }
    @{ id=21; cat="excel"; name="Add status column"; dur=@(5,12); weight=3 }
    @{ id=22; cat="excel"; name="Enter placement type"; dur=@(8,20); weight=2 }
    @{ id=23; cat="excel"; name="Type portfolio name"; dur=@(10,25); weight=2 }
    @{ id=24; cat="excel"; name="Add ASIN value"; dur=@(10,20); weight=3 }
    @{ id=25; cat="excel"; name="Enter negative KW"; dur=@(12,30); weight=3 }

    # === EXCEL - NAVIGATION (20 tasks) ===
    @{ id=26; cat="excel"; name="Scroll down slowly"; dur=@(3,10); weight=5 }
    @{ id=27; cat="excel"; name="Scroll up slowly"; dur=@(3,10); weight=4 }
    @{ id=28; cat="excel"; name="Click random cell"; dur=@(2,6); weight=6 }
    @{ id=29; cat="excel"; name="Select cell range"; dur=@(5,15); weight=4 }
    @{ id=30; cat="excel"; name="Double-click cell"; dur=@(3,8); weight=3 }
    @{ id=31; cat="excel"; name="Press Tab key"; dur=@(1,3); weight=5 }
    @{ id=32; cat="excel"; name="Press Enter key"; dur=@(1,3); weight=5 }
    @{ id=33; cat="excel"; name="Navigate with arrows"; dur=@(3,12); weight=4 }
    @{ id=34; cat="excel"; name="Go to cell Ctrl+G"; dur=@(8,20); weight=2 }
    @{ id=35; cat="excel"; name="Find Ctrl+F search"; dur=@(10,30); weight=3 }
    @{ id=36; cat="excel"; name="Switch worksheet tab"; dur=@(3,8); weight=4 }
    @{ id=37; cat="excel"; name="Resize column width"; dur=@(5,15); weight=2 }
    @{ id=38; cat="excel"; name="Sort column A-Z"; dur=@(5,15); weight=3 }
    @{ id=39; cat="excel"; name="Filter dropdown click"; dur=@(5,15); weight=3 }
    @{ id=40; cat="excel"; name="Clear filter"; dur=@(3,8); weight=2 }
    @{ id=41; cat="excel"; name="Copy cells Ctrl+C"; dur=@(3,8); weight=3 }
    @{ id=42; cat="excel"; name="Paste cells Ctrl+V"; dur=@(3,8); weight=3 }
    @{ id=43; cat="excel"; name="Undo action Ctrl+Z"; dur=@(2,5); weight=2 }
    @{ id=44; cat="excel"; name="Save file Ctrl+S"; dur=@(2,5); weight=4 }
    @{ id=45; cat="excel"; name="Zoom in/out"; dur=@(3,10); weight=2 }

    # === EXCEL - BULK OPERATIONS (10 tasks) ===
    @{ id=46; cat="excel"; name="Review bid column"; dur=@(30,90); weight=3 }
    @{ id=47; cat="excel"; name="Update multiple bids"; dur=@(60,180); weight=2 }
    @{ id=48; cat="excel"; name="Mark negatives batch"; dur=@(45,120); weight=2 }
    @{ id=49; cat="excel"; name="Validate ACOS range"; dur=@(30,90); weight=2 }
    @{ id=50; cat="excel"; name="Check spend totals"; dur=@(20,60); weight=2 }
    @{ id=51; cat="excel"; name="Compare week data"; dur=@(45,120); weight=2 }
    @{ id=52; cat="excel"; name="Highlight outliers"; dur=@(30,90); weight=2 }
    @{ id=53; cat="excel"; name="Format cells batch"; dur=@(20,60); weight=2 }
    @{ id=54; cat="excel"; name="Add conditional format"; dur=@(25,75); weight=2 }
    @{ id=55; cat="excel"; name="Create pivot selection"; dur=@(30,90); weight=1 }

    # === EXCEL - REAL BULK FILE OPERATIONS (15 NEW tasks) ===
    @{ id=131; cat="bulk"; name="Jump to row 50000"; dur=@(8,20); weight=3 }
    @{ id=132; cat="bulk"; name="Jump to row 100000"; dur=@(10,25); weight=2 }
    @{ id=133; cat="bulk"; name="Scroll through bulk data"; dur=@(15,45); weight=4 }
    @{ id=134; cat="bulk"; name="Filter ACOS > 20%"; dur=@(10,30); weight=3 }
    @{ id=135; cat="bulk"; name="Filter by campaign name"; dur=@(10,30); weight=3 }
    @{ id=136; cat="bulk"; name="Sort by Spend DESC"; dur=@(8,25); weight=3 }
    @{ id=137; cat="bulk"; name="Sort by ACOS ASC"; dur=@(8,25); weight=3 }
    @{ id=138; cat="bulk"; name="Clear all filters"; dur=@(5,15); weight=2 }
    @{ id=139; cat="bulk"; name="Navigate to Sheet 2"; dur=@(5,15); weight=2 }
    @{ id=140; cat="bulk"; name="Search specific SKU"; dur=@(15,40); weight=3 }
    @{ id=141; cat="bulk"; name="Review data range"; dur=@(30,90); weight=3 }
    @{ id=142; cat="bulk"; name="Select large range"; dur=@(10,30); weight=2 }
    @{ id=143; cat="bulk"; name="Copy bulk data range"; dur=@(8,20); weight=2 }
    @{ id=144; cat="bulk"; name="Scroll to end of data"; dur=@(8,20); weight=2 }
    @{ id=145; cat="bulk"; name="Return to top of file"; dur=@(5,15); weight=2 }

    # === CHROME - PERPETUA GOALS (20 tasks) ===
    @{ id=56; cat="perpetua"; name="Navigate SP Goals"; dur=@(8,20); weight=4 }
    @{ id=57; cat="perpetua"; name="Click goal row"; dur=@(3,10); weight=5 }
    @{ id=58; cat="perpetua"; name="Review goal ACOS metric"; dur=@(15,45); weight=4 }
    @{ id=59; cat="perpetua"; name="Review goal spend metric"; dur=@(10,30); weight=4 }
    @{ id=60; cat="perpetua"; name="Scroll goals list down"; dur=@(5,15); weight=5 }
    @{ id=61; cat="perpetua"; name="Scroll goals list up"; dur=@(5,15); weight=4 }
    @{ id=62; cat="perpetua"; name="Click goal tabs Keywords"; dur=@(5,15); weight=3 }
    @{ id=63; cat="perpetua"; name="Click goal tabs Negatives"; dur=@(5,15); weight=3 }
    @{ id=64; cat="perpetua"; name="Click goal tabs History"; dur=@(5,15); weight=3 }
    @{ id=65; cat="perpetua"; name="Search goal by name"; dur=@(12,35); weight=3 }
    @{ id=66; cat="perpetua"; name="Filter goals Enabled"; dur=@(5,15); weight=3 }
    @{ id=67; cat="perpetua"; name="Filter goals Paused"; dur=@(5,15); weight=2 }
    @{ id=68; cat="perpetua"; name="Sort by ACOS column"; dur=@(3,10); weight=3 }
    @{ id=69; cat="perpetua"; name="Sort by Spend column"; dur=@(3,10); weight=3 }
    @{ id=70; cat="perpetua"; name="Change date range 7d"; dur=@(8,20); weight=3 }
    @{ id=71; cat="perpetua"; name="Change date range 30d"; dur=@(8,20); weight=3 }
    @{ id=72; cat="perpetua"; name="Click New Goal button"; dur=@(3,8); weight=2 }
    @{ id=73; cat="perpetua"; name="Close goal detail modal"; dur=@(2,5); weight=3 }
    @{ id=74; cat="perpetua"; name="Navigate SB Goals"; dur=@(8,20); weight=3 }
    @{ id=75; cat="perpetua"; name="Navigate SD Goals"; dur=@(8,20); weight=2 }

    # === CHROME - PERPETUA STREAMS (12 tasks) ===
    @{ id=76; cat="perpetua"; name="Navigate SP Streams"; dur=@(8,20); weight=4 }
    @{ id=77; cat="perpetua"; name="Click stream row"; dur=@(3,10); weight=4 }
    @{ id=78; cat="perpetua"; name="Scroll streams list"; dur=@(5,15); weight=4 }
    @{ id=79; cat="perpetua"; name="View stream bid changes"; dur=@(15,40); weight=3 }
    @{ id=80; cat="perpetua"; name="Filter streams by status"; dur=@(5,15); weight=3 }
    @{ id=81; cat="perpetua"; name="Check stream ACOS trend"; dur=@(10,30); weight=3 }
    @{ id=82; cat="perpetua"; name="Review stream keywords"; dur=@(15,45); weight=3 }
    @{ id=83; cat="perpetua"; name="Expand stream details"; dur=@(5,15); weight=3 }
    @{ id=84; cat="perpetua"; name="Collapse stream details"; dur=@(3,8); weight=2 }
    @{ id=85; cat="perpetua"; name="Sort streams by spend"; dur=@(3,10); weight=3 }
    @{ id=86; cat="perpetua"; name="Check stream automation"; dur=@(8,25); weight=2 }
    @{ id=87; cat="perpetua"; name="Navigate Analytics page"; dur=@(8,20); weight=2 }

    # === CHROME - PERPETUA GENERAL (8 tasks) ===
    @{ id=88; cat="perpetua"; name="Click sidebar nav item"; dur=@(3,10); weight=4 }
    @{ id=89; cat="perpetua"; name="Hover sidebar expand"; dur=@(2,6); weight=3 }
    @{ id=90; cat="perpetua"; name="Click account dropdown"; dur=@(3,8); weight=2 }
    @{ id=91; cat="perpetua"; name="Check notifications bell"; dur=@(5,15); weight=3 }
    @{ id=92; cat="perpetua"; name="Refresh current page"; dur=@(3,8); weight=3 }
    @{ id=93; cat="perpetua"; name="Click breadcrumb back"; dur=@(3,8); weight=3 }
    @{ id=94; cat="perpetua"; name="Scroll page randomly"; dur=@(5,20); weight=5 }
    @{ id=95; cat="perpetua"; name="Mouse idle on metrics"; dur=@(8,25); weight=4 }

    # === CHROME - AMAZON ADS (12 tasks) ===
    @{ id=96;  cat="chrome"; name="Load campaign mgr"; dur=@(15,45); weight=3 }
    @{ id=97;  cat="chrome"; name="Click campaign row"; dur=@(5,15); weight=3 }
    @{ id=98;  cat="chrome"; name="View ad group"; dur=@(10,30); weight=3 }
    @{ id=99;  cat="chrome"; name="Check keyword tab"; dur=@(10,30); weight=3 }
    @{ id=100; cat="chrome"; name="Review search terms"; dur=@(20,60); weight=3 }
    @{ id=101; cat="chrome"; name="Scroll campaign list"; dur=@(5,20); weight=3 }
    @{ id=102; cat="chrome"; name="Filter by state"; dur=@(5,15); weight=2 }
    @{ id=103; cat="chrome"; name="Download report"; dur=@(10,30); weight=2 }
    @{ id=104; cat="chrome"; name="Set date picker"; dur=@(10,30); weight=3 }
    @{ id=105; cat="chrome"; name="Check budget status"; dur=@(8,25); weight=3 }
    @{ id=106; cat="chrome"; name="Export to Excel"; dur=@(8,25); weight=2 }
    @{ id=107; cat="chrome"; name="Switch marketplace"; dur=@(8,25); weight=2 }

    # === CHROME - RESEARCH (8 tasks) ===
    @{ id=108; cat="chrome"; name="Google search query"; dur=@(15,45); weight=3 }
    @{ id=109; cat="chrome"; name="Read search result"; dur=@(30,90); weight=3 }
    @{ id=110; cat="chrome"; name="Open new tab"; dur=@(3,8); weight=4 }
    @{ id=111; cat="chrome"; name="Close tab"; dur=@(2,5); weight=3 }
    @{ id=112; cat="chrome"; name="Switch tab"; dur=@(2,5); weight=4 }
    @{ id=113; cat="chrome"; name="Scroll article"; dur=@(10,40); weight=3 }
    @{ id=114; cat="chrome"; name="Click back button"; dur=@(2,5); weight=3 }
    @{ id=115; cat="chrome"; name="Type in address bar"; dur=@(10,30); weight=3 }

    # === TEAMS (10 tasks) ===
    @{ id=116; cat="teams"; name="Click chat thread"; dur=@(5,15); weight=4 }
    @{ id=117; cat="teams"; name="Read message"; dur=@(10,40); weight=4 }
    @{ id=118; cat="teams"; name="Scroll chat history"; dur=@(5,20); weight=3 }
    @{ id=119; cat="teams"; name="Check activity feed"; dur=@(8,25); weight=3 }
    @{ id=120; cat="teams"; name="Click channel"; dur=@(5,15); weight=2 }
    @{ id=121; cat="teams"; name="View files tab"; dur=@(8,25); weight=2 }
    @{ id=122; cat="teams"; name="Search messages"; dur=@(10,30); weight=2 }
    @{ id=123; cat="teams"; name="Check mentions"; dur=@(5,15); weight=3 }
    @{ id=124; cat="teams"; name="Scroll channel list"; dur=@(5,15); weight=2 }
    @{ id=125; cat="teams"; name="React to message"; dur=@(3,8); weight=2 }

    # === HUMAN BEHAVIORS (5 tasks) ===
    @{ id=126; cat="human"; name="Pause and think"; dur=@(5,20); weight=5 }
    @{ id=127; cat="human"; name="Small fidget"; dur=@(2,8); weight=6 }
    @{ id=128; cat="human"; name="Hesitation pause"; dur=@(3,12); weight=4 }
    @{ id=129; cat="human"; name="Re-read and check"; dur=@(8,25); weight=3 }
    @{ id=130; cat="human"; name="Micro-break stretch"; dur=@(15,45); weight=2 }
)

# Category weights (time-of-day adjusted)
function Get-CategoryWeights {
    $hour = (Get-Date).Hour
    if ($hour -ge 9 -and $hour -lt 12) {
        # Morning: more reporting/review - heavy Perpetua + bulk file review
        return @{ excel=25; bulk=15; perpetua=30; chrome=15; teams=10; human=5 }
    } elseif ($hour -ge 12 -and $hour -lt 14) {
    # === EXCEL - BULK FILE OPERATIONS (500+ NEW TASKS) ===
    # Row navigation tasks (100 tasks - jump to specific rows)
    @{ id=146; cat="bulk"; name="Navigate to row 2000"; dur=@(3,8); weight=3 }
    @{ id=147; cat="bulk"; name="Navigate to row 3000"; dur=@(3,8); weight=3 }
    @{ id=148; cat="bulk"; name="Navigate to row 4000"; dur=@(3,8); weight=3 }
    @{ id=149; cat="bulk"; name="Navigate to row 5000"; dur=@(3,8); weight=3 }
    @{ id=150; cat="bulk"; name="Navigate to row 6000"; dur=@(3,8); weight=3 }
    @{ id=151; cat="bulk"; name="Navigate to row 7000"; dur=@(3,8); weight=3 }
    @{ id=152; cat="bulk"; name="Navigate to row 8000"; dur=@(3,8); weight=3 }
    @{ id=153; cat="bulk"; name="Navigate to row 9000"; dur=@(3,8); weight=3 }
    @{ id=154; cat="bulk"; name="Navigate to row 10000"; dur=@(3,8); weight=3 }
    @{ id=155; cat="bulk"; name="Navigate to row 11000"; dur=@(3,8); weight=3 }
    @{ id=156; cat="bulk"; name="Navigate to row 12000"; dur=@(3,8); weight=3 }
    @{ id=157; cat="bulk"; name="Navigate to row 13000"; dur=@(3,8); weight=3 }
    @{ id=158; cat="bulk"; name="Navigate to row 14000"; dur=@(3,8); weight=3 }
    @{ id=159; cat="bulk"; name="Navigate to row 15000"; dur=@(3,8); weight=3 }
    @{ id=160; cat="bulk"; name="Navigate to row 16000"; dur=@(3,8); weight=3 }
    @{ id=161; cat="bulk"; name="Navigate to row 17000"; dur=@(3,8); weight=3 }
    @{ id=162; cat="bulk"; name="Navigate to row 18000"; dur=@(3,8); weight=3 }
    @{ id=163; cat="bulk"; name="Navigate to row 19000"; dur=@(3,8); weight=3 }
    @{ id=164; cat="bulk"; name="Navigate to row 20000"; dur=@(3,8); weight=3 }
    @{ id=165; cat="bulk"; name="Navigate to row 21000"; dur=@(3,8); weight=3 }
    @{ id=166; cat="bulk"; name="Navigate to row 22000"; dur=@(3,8); weight=3 }
    @{ id=167; cat="bulk"; name="Navigate to row 23000"; dur=@(3,8); weight=3 }
    @{ id=168; cat="bulk"; name="Navigate to row 24000"; dur=@(3,8); weight=3 }
    @{ id=169; cat="bulk"; name="Navigate to row 25000"; dur=@(3,8); weight=3 }
    @{ id=170; cat="bulk"; name="Navigate to row 26000"; dur=@(3,8); weight=3 }
    @{ id=171; cat="bulk"; name="Navigate to row 27000"; dur=@(3,8); weight=3 }
    @{ id=172; cat="bulk"; name="Navigate to row 28000"; dur=@(3,8); weight=3 }
    @{ id=173; cat="bulk"; name="Navigate to row 29000"; dur=@(3,8); weight=3 }
    @{ id=174; cat="bulk"; name="Navigate to row 30000"; dur=@(3,8); weight=3 }
    @{ id=175; cat="bulk"; name="Navigate to row 31000"; dur=@(3,8); weight=3 }
    @{ id=176; cat="bulk"; name="Navigate to row 32000"; dur=@(3,8); weight=3 }
    @{ id=177; cat="bulk"; name="Navigate to row 33000"; dur=@(3,8); weight=3 }
    @{ id=178; cat="bulk"; name="Navigate to row 34000"; dur=@(3,8); weight=3 }
    @{ id=179; cat="bulk"; name="Navigate to row 35000"; dur=@(3,8); weight=3 }
    @{ id=180; cat="bulk"; name="Navigate to row 36000"; dur=@(3,8); weight=3 }
    @{ id=181; cat="bulk"; name="Navigate to row 37000"; dur=@(3,8); weight=3 }
    @{ id=182; cat="bulk"; name="Navigate to row 38000"; dur=@(3,8); weight=3 }
    @{ id=183; cat="bulk"; name="Navigate to row 39000"; dur=@(3,8); weight=3 }
    @{ id=184; cat="bulk"; name="Navigate to row 40000"; dur=@(3,8); weight=3 }
    @{ id=185; cat="bulk"; name="Navigate to row 41000"; dur=@(3,8); weight=3 }
    @{ id=186; cat="bulk"; name="Navigate to row 42000"; dur=@(3,8); weight=3 }
    @{ id=187; cat="bulk"; name="Navigate to row 43000"; dur=@(3,8); weight=3 }
    @{ id=188; cat="bulk"; name="Navigate to row 44000"; dur=@(3,8); weight=3 }
    @{ id=189; cat="bulk"; name="Navigate to row 45000"; dur=@(3,8); weight=3 }
    @{ id=190; cat="bulk"; name="Navigate to row 46000"; dur=@(3,8); weight=3 }
    @{ id=191; cat="bulk"; name="Navigate to row 47000"; dur=@(3,8); weight=3 }
    @{ id=192; cat="bulk"; name="Navigate to row 48000"; dur=@(3,8); weight=3 }
    @{ id=193; cat="bulk"; name="Navigate to row 49000"; dur=@(3,8); weight=3 }
    @{ id=194; cat="bulk"; name="Navigate to row 50000"; dur=@(3,8); weight=3 }
    @{ id=195; cat="bulk"; name="Navigate to row 51000"; dur=@(3,8); weight=3 }
    @{ id=196; cat="bulk"; name="Navigate to row 52000"; dur=@(3,8); weight=3 }
    @{ id=197; cat="bulk"; name="Navigate to row 53000"; dur=@(3,8); weight=3 }
    @{ id=198; cat="bulk"; name="Navigate to row 54000"; dur=@(3,8); weight=3 }
    @{ id=199; cat="bulk"; name="Navigate to row 55000"; dur=@(3,8); weight=3 }
    @{ id=200; cat="bulk"; name="Navigate to row 56000"; dur=@(3,8); weight=3 }
    @{ id=201; cat="bulk"; name="Navigate to row 57000"; dur=@(3,8); weight=3 }
    @{ id=202; cat="bulk"; name="Navigate to row 58000"; dur=@(3,8); weight=3 }
    @{ id=203; cat="bulk"; name="Navigate to row 59000"; dur=@(3,8); weight=3 }
    @{ id=204; cat="bulk"; name="Navigate to row 60000"; dur=@(3,8); weight=3 }
    @{ id=205; cat="bulk"; name="Navigate to row 61000"; dur=@(3,8); weight=3 }
    @{ id=206; cat="bulk"; name="Navigate to row 62000"; dur=@(3,8); weight=3 }
    @{ id=207; cat="bulk"; name="Navigate to row 63000"; dur=@(3,8); weight=3 }
    @{ id=208; cat="bulk"; name="Navigate to row 64000"; dur=@(3,8); weight=3 }
    @{ id=209; cat="bulk"; name="Navigate to row 65000"; dur=@(3,8); weight=3 }
    @{ id=210; cat="bulk"; name="Navigate to row 66000"; dur=@(3,8); weight=3 }
    @{ id=211; cat="bulk"; name="Navigate to row 67000"; dur=@(3,8); weight=3 }
    @{ id=212; cat="bulk"; name="Navigate to row 68000"; dur=@(3,8); weight=3 }
    @{ id=213; cat="bulk"; name="Navigate to row 69000"; dur=@(3,8); weight=3 }
    @{ id=214; cat="bulk"; name="Navigate to row 70000"; dur=@(3,8); weight=3 }
    @{ id=215; cat="bulk"; name="Navigate to row 71000"; dur=@(3,8); weight=3 }
    @{ id=216; cat="bulk"; name="Navigate to row 72000"; dur=@(3,8); weight=3 }
    @{ id=217; cat="bulk"; name="Navigate to row 73000"; dur=@(3,8); weight=3 }
    @{ id=218; cat="bulk"; name="Navigate to row 74000"; dur=@(3,8); weight=3 }
    @{ id=219; cat="bulk"; name="Navigate to row 75000"; dur=@(3,8); weight=3 }
    @{ id=220; cat="bulk"; name="Navigate to row 76000"; dur=@(3,8); weight=3 }
    @{ id=221; cat="bulk"; name="Navigate to row 77000"; dur=@(3,8); weight=3 }
    @{ id=222; cat="bulk"; name="Navigate to row 78000"; dur=@(3,8); weight=3 }
    @{ id=223; cat="bulk"; name="Navigate to row 79000"; dur=@(3,8); weight=3 }
    @{ id=224; cat="bulk"; name="Navigate to row 80000"; dur=@(3,8); weight=3 }
    @{ id=225; cat="bulk"; name="Navigate to row 81000"; dur=@(3,8); weight=3 }
    @{ id=226; cat="bulk"; name="Navigate to row 82000"; dur=@(3,8); weight=3 }
    @{ id=227; cat="bulk"; name="Navigate to row 83000"; dur=@(3,8); weight=3 }
    @{ id=228; cat="bulk"; name="Navigate to row 84000"; dur=@(3,8); weight=3 }
    @{ id=229; cat="bulk"; name="Navigate to row 85000"; dur=@(3,8); weight=3 }
    @{ id=230; cat="bulk"; name="Navigate to row 86000"; dur=@(3,8); weight=3 }
    @{ id=231; cat="bulk"; name="Navigate to row 87000"; dur=@(3,8); weight=3 }
    @{ id=232; cat="bulk"; name="Navigate to row 88000"; dur=@(3,8); weight=3 }
    @{ id=233; cat="bulk"; name="Navigate to row 89000"; dur=@(3,8); weight=3 }
    @{ id=234; cat="bulk"; name="Navigate to row 90000"; dur=@(3,8); weight=3 }
    @{ id=235; cat="bulk"; name="Navigate to row 91000"; dur=@(3,8); weight=3 }
    @{ id=236; cat="bulk"; name="Navigate to row 92000"; dur=@(3,8); weight=3 }
    @{ id=237; cat="bulk"; name="Navigate to row 93000"; dur=@(3,8); weight=3 }
    @{ id=238; cat="bulk"; name="Navigate to row 94000"; dur=@(3,8); weight=3 }
    @{ id=239; cat="bulk"; name="Navigate to row 95000"; dur=@(3,8); weight=3 }
    @{ id=240; cat="bulk"; name="Navigate to row 96000"; dur=@(3,8); weight=3 }
    @{ id=241; cat="bulk"; name="Navigate to row 97000"; dur=@(3,8); weight=3 }
    @{ id=242; cat="bulk"; name="Navigate to row 98000"; dur=@(3,8); weight=3 }
    @{ id=243; cat="bulk"; name="Navigate to row 99000"; dur=@(3,8); weight=3 }
    @{ id=244; cat="bulk"; name="Navigate to row 100000"; dur=@(3,8); weight=3 }
    @{ id=245; cat="bulk"; name="Navigate to row 101000"; dur=@(3,8); weight=3 }
    @{ id=246; cat="bulk"; name="Filter ACOS > 12%"; dur=@(4,10); weight=2 }
    @{ id=247; cat="bulk"; name="Filter ACOS > 14%"; dur=@(4,10); weight=2 }
    @{ id=248; cat="bulk"; name="Filter ACOS > 16%"; dur=@(4,10); weight=2 }
    @{ id=249; cat="bulk"; name="Filter ACOS > 18%"; dur=@(4,10); weight=2 }
    @{ id=250; cat="bulk"; name="Filter ACOS > 20%"; dur=@(4,10); weight=2 }
    @{ id=251; cat="bulk"; name="Filter ACOS > 22%"; dur=@(4,10); weight=2 }
    @{ id=252; cat="bulk"; name="Filter ACOS > 24%"; dur=@(4,10); weight=2 }
    @{ id=253; cat="bulk"; name="Filter ACOS > 26%"; dur=@(4,10); weight=2 }
    @{ id=254; cat="bulk"; name="Filter ACOS > 28%"; dur=@(4,10); weight=2 }
    @{ id=255; cat="bulk"; name="Filter ACOS > 30%"; dur=@(4,10); weight=2 }
    @{ id=256; cat="bulk"; name="Filter ACOS > 32%"; dur=@(4,10); weight=2 }
    @{ id=257; cat="bulk"; name="Filter ACOS > 34%"; dur=@(4,10); weight=2 }
    @{ id=258; cat="bulk"; name="Filter ACOS > 36%"; dur=@(4,10); weight=2 }
    @{ id=259; cat="bulk"; name="Filter ACOS > 38%"; dur=@(4,10); weight=2 }
    @{ id=260; cat="bulk"; name="Filter ACOS > 40%"; dur=@(4,10); weight=2 }
    @{ id=261; cat="bulk"; name="Filter ACOS > 42%"; dur=@(4,10); weight=2 }
    @{ id=262; cat="bulk"; name="Filter ACOS > 44%"; dur=@(4,10); weight=2 }
    @{ id=263; cat="bulk"; name="Filter ACOS > 46%"; dur=@(4,10); weight=2 }
    @{ id=264; cat="bulk"; name="Filter ACOS > 48%"; dur=@(4,10); weight=2 }
    @{ id=265; cat="bulk"; name="Filter ACOS > 50%"; dur=@(4,10); weight=2 }
    @{ id=266; cat="bulk"; name="Filter ACOS > 52%"; dur=@(4,10); weight=2 }
    @{ id=267; cat="bulk"; name="Filter ACOS > 54%"; dur=@(4,10); weight=2 }
    @{ id=268; cat="bulk"; name="Filter ACOS > 56%"; dur=@(4,10); weight=2 }
    @{ id=269; cat="bulk"; name="Filter ACOS > 58%"; dur=@(4,10); weight=2 }
    @{ id=270; cat="bulk"; name="Filter ACOS > 60%"; dur=@(4,10); weight=2 }
    @{ id=271; cat="bulk"; name="Filter ACOS > 62%"; dur=@(4,10); weight=2 }
    @{ id=272; cat="bulk"; name="Filter ACOS > 64%"; dur=@(4,10); weight=2 }
    @{ id=273; cat="bulk"; name="Filter ACOS > 66%"; dur=@(4,10); weight=2 }
    @{ id=274; cat="bulk"; name="Filter ACOS > 68%"; dur=@(4,10); weight=2 }
    @{ id=275; cat="bulk"; name="Filter ACOS > 70%"; dur=@(4,10); weight=2 }
    @{ id=276; cat="bulk"; name="Filter ACOS > 72%"; dur=@(4,10); weight=2 }
    @{ id=277; cat="bulk"; name="Filter ACOS > 74%"; dur=@(4,10); weight=2 }
    @{ id=278; cat="bulk"; name="Filter ACOS > 76%"; dur=@(4,10); weight=2 }
    @{ id=279; cat="bulk"; name="Filter ACOS > 78%"; dur=@(4,10); weight=2 }
    @{ id=280; cat="bulk"; name="Filter ACOS > 80%"; dur=@(4,10); weight=2 }
    @{ id=281; cat="bulk"; name="Filter ACOS > 82%"; dur=@(4,10); weight=2 }
    @{ id=282; cat="bulk"; name="Filter ACOS > 84%"; dur=@(4,10); weight=2 }
    @{ id=283; cat="bulk"; name="Filter ACOS > 86%"; dur=@(4,10); weight=2 }
    @{ id=284; cat="bulk"; name="Filter ACOS > 88%"; dur=@(4,10); weight=2 }
    @{ id=285; cat="bulk"; name="Filter ACOS > 90%"; dur=@(4,10); weight=2 }
    @{ id=286; cat="bulk"; name="Filter ACOS > 92%"; dur=@(4,10); weight=2 }
    @{ id=287; cat="bulk"; name="Filter ACOS > 94%"; dur=@(4,10); weight=2 }
    @{ id=288; cat="bulk"; name="Filter ACOS > 96%"; dur=@(4,10); weight=2 }
    @{ id=289; cat="bulk"; name="Filter ACOS > 98%"; dur=@(4,10); weight=2 }
    @{ id=290; cat="bulk"; name="Filter ACOS > 100%"; dur=@(4,10); weight=2 }
    @{ id=291; cat="bulk"; name="Filter ACOS > 102%"; dur=@(4,10); weight=2 }
    @{ id=292; cat="bulk"; name="Filter ACOS > 104%"; dur=@(4,10); weight=2 }
    @{ id=293; cat="bulk"; name="Filter ACOS > 106%"; dur=@(4,10); weight=2 }
    @{ id=294; cat="bulk"; name="Filter ACOS > 108%"; dur=@(4,10); weight=2 }
    @{ id=295; cat="bulk"; name="Filter ACOS > 110%"; dur=@(4,10); weight=2 }
    @{ id=296; cat="bulk"; name="Filter ACOS > 112%"; dur=@(4,10); weight=2 }
    @{ id=297; cat="bulk"; name="Filter ACOS > 114%"; dur=@(4,10); weight=2 }
    @{ id=298; cat="bulk"; name="Filter ACOS > 116%"; dur=@(4,10); weight=2 }
    @{ id=299; cat="bulk"; name="Filter ACOS > 118%"; dur=@(4,10); weight=2 }
    @{ id=300; cat="bulk"; name="Filter ACOS > 120%"; dur=@(4,10); weight=2 }
    @{ id=301; cat="bulk"; name="Filter ACOS > 122%"; dur=@(4,10); weight=2 }
    @{ id=302; cat="bulk"; name="Filter ACOS > 124%"; dur=@(4,10); weight=2 }
    @{ id=303; cat="bulk"; name="Filter ACOS > 126%"; dur=@(4,10); weight=2 }
    @{ id=304; cat="bulk"; name="Filter ACOS > 128%"; dur=@(4,10); weight=2 }
    @{ id=305; cat="bulk"; name="Filter ACOS > 130%"; dur=@(4,10); weight=2 }
    @{ id=306; cat="bulk"; name="Filter ACOS > 132%"; dur=@(4,10); weight=2 }
    @{ id=307; cat="bulk"; name="Filter ACOS > 134%"; dur=@(4,10); weight=2 }
    @{ id=308; cat="bulk"; name="Filter ACOS > 136%"; dur=@(4,10); weight=2 }
    @{ id=309; cat="bulk"; name="Filter ACOS > 138%"; dur=@(4,10); weight=2 }
    @{ id=310; cat="bulk"; name="Filter ACOS > 140%"; dur=@(4,10); weight=2 }
    @{ id=311; cat="bulk"; name="Filter ACOS > 142%"; dur=@(4,10); weight=2 }
    @{ id=312; cat="bulk"; name="Filter ACOS > 144%"; dur=@(4,10); weight=2 }
    @{ id=313; cat="bulk"; name="Filter ACOS > 146%"; dur=@(4,10); weight=2 }
    @{ id=314; cat="bulk"; name="Filter ACOS > 148%"; dur=@(4,10); weight=2 }
    @{ id=315; cat="bulk"; name="Filter ACOS > 150%"; dur=@(4,10); weight=2 }
    @{ id=316; cat="bulk"; name="Filter ACOS > 152%"; dur=@(4,10); weight=2 }
    @{ id=317; cat="bulk"; name="Filter ACOS > 154%"; dur=@(4,10); weight=2 }
    @{ id=318; cat="bulk"; name="Filter ACOS > 156%"; dur=@(4,10); weight=2 }
    @{ id=319; cat="bulk"; name="Filter ACOS > 158%"; dur=@(4,10); weight=2 }
    @{ id=320; cat="bulk"; name="Filter ACOS > 160%"; dur=@(4,10); weight=2 }
    @{ id=321; cat="bulk"; name="Filter ACOS > 162%"; dur=@(4,10); weight=2 }
    @{ id=322; cat="bulk"; name="Filter ACOS > 164%"; dur=@(4,10); weight=2 }
    @{ id=323; cat="bulk"; name="Filter ACOS > 166%"; dur=@(4,10); weight=2 }
    @{ id=324; cat="bulk"; name="Filter ACOS > 168%"; dur=@(4,10); weight=2 }
    @{ id=325; cat="bulk"; name="Filter ACOS > 170%"; dur=@(4,10); weight=2 }
    @{ id=326; cat="bulk"; name="Filter ACOS > 172%"; dur=@(4,10); weight=2 }
    @{ id=327; cat="bulk"; name="Filter ACOS > 174%"; dur=@(4,10); weight=2 }
    @{ id=328; cat="bulk"; name="Filter ACOS > 176%"; dur=@(4,10); weight=2 }
    @{ id=329; cat="bulk"; name="Filter ACOS > 178%"; dur=@(4,10); weight=2 }
    @{ id=330; cat="bulk"; name="Filter ACOS > 180%"; dur=@(4,10); weight=2 }
    @{ id=331; cat="bulk"; name="Filter ACOS > 182%"; dur=@(4,10); weight=2 }
    @{ id=332; cat="bulk"; name="Filter ACOS > 184%"; dur=@(4,10); weight=2 }
    @{ id=333; cat="bulk"; name="Filter ACOS > 186%"; dur=@(4,10); weight=2 }
    @{ id=334; cat="bulk"; name="Filter ACOS > 188%"; dur=@(4,10); weight=2 }
    @{ id=335; cat="bulk"; name="Filter ACOS > 190%"; dur=@(4,10); weight=2 }
    @{ id=336; cat="bulk"; name="Filter ACOS > 192%"; dur=@(4,10); weight=2 }
    @{ id=337; cat="bulk"; name="Filter ACOS > 194%"; dur=@(4,10); weight=2 }
    @{ id=338; cat="bulk"; name="Filter ACOS > 196%"; dur=@(4,10); weight=2 }
    @{ id=339; cat="bulk"; name="Filter ACOS > 198%"; dur=@(4,10); weight=2 }
    @{ id=340; cat="bulk"; name="Filter ACOS > 200%"; dur=@(4,10); weight=2 }
    @{ id=341; cat="bulk"; name="Filter ACOS > 202%"; dur=@(4,10); weight=2 }
    @{ id=342; cat="bulk"; name="Filter ACOS > 204%"; dur=@(4,10); weight=2 }
    @{ id=343; cat="bulk"; name="Filter ACOS > 206%"; dur=@(4,10); weight=2 }
    @{ id=344; cat="bulk"; name="Filter ACOS > 208%"; dur=@(4,10); weight=2 }
    @{ id=345; cat="bulk"; name="Filter ACOS > 210%"; dur=@(4,10); weight=2 }
    @{ id=346; cat="bulk"; name="Sort by ACOS column"; dur=@(3,8); weight=2 }
    @{ id=347; cat="bulk"; name="Sort by Sales column"; dur=@(3,8); weight=2 }
    @{ id=348; cat="bulk"; name="Sort by Clicks column"; dur=@(3,8); weight=2 }
    @{ id=349; cat="bulk"; name="Sort by Impressions column"; dur=@(3,8); weight=2 }
    @{ id=350; cat="bulk"; name="Sort by CTR column"; dur=@(3,8); weight=2 }
    @{ id=351; cat="bulk"; name="Sort by CPC column"; dur=@(3,8); weight=2 }
    @{ id=352; cat="bulk"; name="Sort by ROAS column"; dur=@(3,8); weight=2 }
    @{ id=353; cat="bulk"; name="Sort by Spend column"; dur=@(3,8); weight=2 }
    @{ id=354; cat="bulk"; name="Sort by ACOS column"; dur=@(3,8); weight=2 }
    @{ id=355; cat="bulk"; name="Sort by Sales column"; dur=@(3,8); weight=2 }
    @{ id=356; cat="bulk"; name="Sort by Clicks column"; dur=@(3,8); weight=2 }
    @{ id=357; cat="bulk"; name="Sort by Impressions column"; dur=@(3,8); weight=2 }
    @{ id=358; cat="bulk"; name="Sort by CTR column"; dur=@(3,8); weight=2 }
    @{ id=359; cat="bulk"; name="Sort by CPC column"; dur=@(3,8); weight=2 }
    @{ id=360; cat="bulk"; name="Sort by ROAS column"; dur=@(3,8); weight=2 }
    @{ id=361; cat="bulk"; name="Sort by Spend column"; dur=@(3,8); weight=2 }
    @{ id=362; cat="bulk"; name="Sort by ACOS column"; dur=@(3,8); weight=2 }
    @{ id=363; cat="bulk"; name="Sort by Sales column"; dur=@(3,8); weight=2 }
    @{ id=364; cat="bulk"; name="Sort by Clicks column"; dur=@(3,8); weight=2 }
    @{ id=365; cat="bulk"; name="Sort by Impressions column"; dur=@(3,8); weight=2 }
    @{ id=366; cat="bulk"; name="Sort by CTR column"; dur=@(3,8); weight=2 }
    @{ id=367; cat="bulk"; name="Sort by CPC column"; dur=@(3,8); weight=2 }
    @{ id=368; cat="bulk"; name="Sort by ROAS column"; dur=@(3,8); weight=2 }
    @{ id=369; cat="bulk"; name="Sort by Spend column"; dur=@(3,8); weight=2 }
    @{ id=370; cat="bulk"; name="Sort by ACOS column"; dur=@(3,8); weight=2 }
    @{ id=371; cat="bulk"; name="Sort by Sales column"; dur=@(3,8); weight=2 }
    @{ id=372; cat="bulk"; name="Sort by Clicks column"; dur=@(3,8); weight=2 }
    @{ id=373; cat="bulk"; name="Sort by Impressions column"; dur=@(3,8); weight=2 }
    @{ id=374; cat="bulk"; name="Sort by CTR column"; dur=@(3,8); weight=2 }
    @{ id=375; cat="bulk"; name="Sort by CPC column"; dur=@(3,8); weight=2 }
    @{ id=376; cat="bulk"; name="Sort by ROAS column"; dur=@(3,8); weight=2 }
    @{ id=377; cat="bulk"; name="Sort by Spend column"; dur=@(3,8); weight=2 }
    @{ id=378; cat="bulk"; name="Sort by ACOS column"; dur=@(3,8); weight=2 }
    @{ id=379; cat="bulk"; name="Sort by Sales column"; dur=@(3,8); weight=2 }
    @{ id=380; cat="bulk"; name="Sort by Clicks column"; dur=@(3,8); weight=2 }
    @{ id=381; cat="bulk"; name="Sort by Impressions column"; dur=@(3,8); weight=2 }
    @{ id=382; cat="bulk"; name="Sort by CTR column"; dur=@(3,8); weight=2 }
    @{ id=383; cat="bulk"; name="Sort by CPC column"; dur=@(3,8); weight=2 }
    @{ id=384; cat="bulk"; name="Sort by ROAS column"; dur=@(3,8); weight=2 }
    @{ id=385; cat="bulk"; name="Sort by Spend column"; dur=@(3,8); weight=2 }
    @{ id=386; cat="bulk"; name="Sort by ACOS column"; dur=@(3,8); weight=2 }
    @{ id=387; cat="bulk"; name="Sort by Sales column"; dur=@(3,8); weight=2 }
    @{ id=388; cat="bulk"; name="Sort by Clicks column"; dur=@(3,8); weight=2 }
    @{ id=389; cat="bulk"; name="Sort by Impressions column"; dur=@(3,8); weight=2 }
    @{ id=390; cat="bulk"; name="Sort by CTR column"; dur=@(3,8); weight=2 }
    @{ id=391; cat="bulk"; name="Sort by CPC column"; dur=@(3,8); weight=2 }
    @{ id=392; cat="bulk"; name="Sort by ROAS column"; dur=@(3,8); weight=2 }
    @{ id=393; cat="bulk"; name="Sort by Spend column"; dur=@(3,8); weight=2 }
    @{ id=394; cat="bulk"; name="Sort by ACOS column"; dur=@(3,8); weight=2 }
    @{ id=395; cat="bulk"; name="Sort by Sales column"; dur=@(3,8); weight=2 }
    @{ id=396; cat="bulk"; name="Scroll rows 500-1000"; dur=@(5,15); weight=3 }
    @{ id=397; cat="bulk"; name="Scroll rows 1000-1500"; dur=@(5,15); weight=3 }
    @{ id=398; cat="bulk"; name="Scroll rows 1500-2000"; dur=@(5,15); weight=3 }
    @{ id=399; cat="bulk"; name="Scroll rows 2000-2500"; dur=@(5,15); weight=3 }
    @{ id=400; cat="bulk"; name="Scroll rows 2500-3000"; dur=@(5,15); weight=3 }
    @{ id=401; cat="bulk"; name="Scroll rows 3000-3500"; dur=@(5,15); weight=3 }
    @{ id=402; cat="bulk"; name="Scroll rows 3500-4000"; dur=@(5,15); weight=3 }
    @{ id=403; cat="bulk"; name="Scroll rows 4000-4500"; dur=@(5,15); weight=3 }
    @{ id=404; cat="bulk"; name="Scroll rows 4500-5000"; dur=@(5,15); weight=3 }
    @{ id=405; cat="bulk"; name="Scroll rows 5000-5500"; dur=@(5,15); weight=3 }
    @{ id=406; cat="bulk"; name="Scroll rows 5500-6000"; dur=@(5,15); weight=3 }
    @{ id=407; cat="bulk"; name="Scroll rows 6000-6500"; dur=@(5,15); weight=3 }
    @{ id=408; cat="bulk"; name="Scroll rows 6500-7000"; dur=@(5,15); weight=3 }
    @{ id=409; cat="bulk"; name="Scroll rows 7000-7500"; dur=@(5,15); weight=3 }
    @{ id=410; cat="bulk"; name="Scroll rows 7500-8000"; dur=@(5,15); weight=3 }
    @{ id=411; cat="bulk"; name="Scroll rows 8000-8500"; dur=@(5,15); weight=3 }
    @{ id=412; cat="bulk"; name="Scroll rows 8500-9000"; dur=@(5,15); weight=3 }
    @{ id=413; cat="bulk"; name="Scroll rows 9000-9500"; dur=@(5,15); weight=3 }
    @{ id=414; cat="bulk"; name="Scroll rows 9500-10000"; dur=@(5,15); weight=3 }
    @{ id=415; cat="bulk"; name="Scroll rows 10000-10500"; dur=@(5,15); weight=3 }
    @{ id=416; cat="bulk"; name="Scroll rows 10500-11000"; dur=@(5,15); weight=3 }
    @{ id=417; cat="bulk"; name="Scroll rows 11000-11500"; dur=@(5,15); weight=3 }
    @{ id=418; cat="bulk"; name="Scroll rows 11500-12000"; dur=@(5,15); weight=3 }
    @{ id=419; cat="bulk"; name="Scroll rows 12000-12500"; dur=@(5,15); weight=3 }
    @{ id=420; cat="bulk"; name="Scroll rows 12500-13000"; dur=@(5,15); weight=3 }
    @{ id=421; cat="bulk"; name="Scroll rows 13000-13500"; dur=@(5,15); weight=3 }
    @{ id=422; cat="bulk"; name="Scroll rows 13500-14000"; dur=@(5,15); weight=3 }
    @{ id=423; cat="bulk"; name="Scroll rows 14000-14500"; dur=@(5,15); weight=3 }
    @{ id=424; cat="bulk"; name="Scroll rows 14500-15000"; dur=@(5,15); weight=3 }
    @{ id=425; cat="bulk"; name="Scroll rows 15000-15500"; dur=@(5,15); weight=3 }
    @{ id=426; cat="bulk"; name="Scroll rows 15500-16000"; dur=@(5,15); weight=3 }
    @{ id=427; cat="bulk"; name="Scroll rows 16000-16500"; dur=@(5,15); weight=3 }
    @{ id=428; cat="bulk"; name="Scroll rows 16500-17000"; dur=@(5,15); weight=3 }
    @{ id=429; cat="bulk"; name="Scroll rows 17000-17500"; dur=@(5,15); weight=3 }
    @{ id=430; cat="bulk"; name="Scroll rows 17500-18000"; dur=@(5,15); weight=3 }
    @{ id=431; cat="bulk"; name="Scroll rows 18000-18500"; dur=@(5,15); weight=3 }
    @{ id=432; cat="bulk"; name="Scroll rows 18500-19000"; dur=@(5,15); weight=3 }
    @{ id=433; cat="bulk"; name="Scroll rows 19000-19500"; dur=@(5,15); weight=3 }
    @{ id=434; cat="bulk"; name="Scroll rows 19500-20000"; dur=@(5,15); weight=3 }
    @{ id=435; cat="bulk"; name="Scroll rows 20000-20500"; dur=@(5,15); weight=3 }
    @{ id=436; cat="bulk"; name="Scroll rows 20500-21000"; dur=@(5,15); weight=3 }
    @{ id=437; cat="bulk"; name="Scroll rows 21000-21500"; dur=@(5,15); weight=3 }
    @{ id=438; cat="bulk"; name="Scroll rows 21500-22000"; dur=@(5,15); weight=3 }
    @{ id=439; cat="bulk"; name="Scroll rows 22000-22500"; dur=@(5,15); weight=3 }
    @{ id=440; cat="bulk"; name="Scroll rows 22500-23000"; dur=@(5,15); weight=3 }
    @{ id=441; cat="bulk"; name="Scroll rows 23000-23500"; dur=@(5,15); weight=3 }
    @{ id=442; cat="bulk"; name="Scroll rows 23500-24000"; dur=@(5,15); weight=3 }
    @{ id=443; cat="bulk"; name="Scroll rows 24000-24500"; dur=@(5,15); weight=3 }
    @{ id=444; cat="bulk"; name="Scroll rows 24500-25000"; dur=@(5,15); weight=3 }
    @{ id=445; cat="bulk"; name="Scroll rows 25000-25500"; dur=@(5,15); weight=3 }
    @{ id=446; cat="bulk"; name="Scroll rows 25500-26000"; dur=@(5,15); weight=3 }
    @{ id=447; cat="bulk"; name="Scroll rows 26000-26500"; dur=@(5,15); weight=3 }
    @{ id=448; cat="bulk"; name="Scroll rows 26500-27000"; dur=@(5,15); weight=3 }
    @{ id=449; cat="bulk"; name="Scroll rows 27000-27500"; dur=@(5,15); weight=3 }
    @{ id=450; cat="bulk"; name="Scroll rows 27500-28000"; dur=@(5,15); weight=3 }
    @{ id=451; cat="bulk"; name="Scroll rows 28000-28500"; dur=@(5,15); weight=3 }
    @{ id=452; cat="bulk"; name="Scroll rows 28500-29000"; dur=@(5,15); weight=3 }
    @{ id=453; cat="bulk"; name="Scroll rows 29000-29500"; dur=@(5,15); weight=3 }
    @{ id=454; cat="bulk"; name="Scroll rows 29500-30000"; dur=@(5,15); weight=3 }
    @{ id=455; cat="bulk"; name="Scroll rows 30000-30500"; dur=@(5,15); weight=3 }
    @{ id=456; cat="bulk"; name="Scroll rows 30500-31000"; dur=@(5,15); weight=3 }
    @{ id=457; cat="bulk"; name="Scroll rows 31000-31500"; dur=@(5,15); weight=3 }
    @{ id=458; cat="bulk"; name="Scroll rows 31500-32000"; dur=@(5,15); weight=3 }
    @{ id=459; cat="bulk"; name="Scroll rows 32000-32500"; dur=@(5,15); weight=3 }
    @{ id=460; cat="bulk"; name="Scroll rows 32500-33000"; dur=@(5,15); weight=3 }
    @{ id=461; cat="bulk"; name="Scroll rows 33000-33500"; dur=@(5,15); weight=3 }
    @{ id=462; cat="bulk"; name="Scroll rows 33500-34000"; dur=@(5,15); weight=3 }
    @{ id=463; cat="bulk"; name="Scroll rows 34000-34500"; dur=@(5,15); weight=3 }
    @{ id=464; cat="bulk"; name="Scroll rows 34500-35000"; dur=@(5,15); weight=3 }
    @{ id=465; cat="bulk"; name="Scroll rows 35000-35500"; dur=@(5,15); weight=3 }
    @{ id=466; cat="bulk"; name="Scroll rows 35500-36000"; dur=@(5,15); weight=3 }
    @{ id=467; cat="bulk"; name="Scroll rows 36000-36500"; dur=@(5,15); weight=3 }
    @{ id=468; cat="bulk"; name="Scroll rows 36500-37000"; dur=@(5,15); weight=3 }
    @{ id=469; cat="bulk"; name="Scroll rows 37000-37500"; dur=@(5,15); weight=3 }
    @{ id=470; cat="bulk"; name="Scroll rows 37500-38000"; dur=@(5,15); weight=3 }
    @{ id=471; cat="bulk"; name="Scroll rows 38000-38500"; dur=@(5,15); weight=3 }
    @{ id=472; cat="bulk"; name="Scroll rows 38500-39000"; dur=@(5,15); weight=3 }
    @{ id=473; cat="bulk"; name="Scroll rows 39000-39500"; dur=@(5,15); weight=3 }
    @{ id=474; cat="bulk"; name="Scroll rows 39500-40000"; dur=@(5,15); weight=3 }
    @{ id=475; cat="bulk"; name="Scroll rows 40000-40500"; dur=@(5,15); weight=3 }
    @{ id=476; cat="bulk"; name="Scroll rows 40500-41000"; dur=@(5,15); weight=3 }
    @{ id=477; cat="bulk"; name="Scroll rows 41000-41500"; dur=@(5,15); weight=3 }
    @{ id=478; cat="bulk"; name="Scroll rows 41500-42000"; dur=@(5,15); weight=3 }
    @{ id=479; cat="bulk"; name="Scroll rows 42000-42500"; dur=@(5,15); weight=3 }
    @{ id=480; cat="bulk"; name="Scroll rows 42500-43000"; dur=@(5,15); weight=3 }
    @{ id=481; cat="bulk"; name="Scroll rows 43000-43500"; dur=@(5,15); weight=3 }
    @{ id=482; cat="bulk"; name="Scroll rows 43500-44000"; dur=@(5,15); weight=3 }
    @{ id=483; cat="bulk"; name="Scroll rows 44000-44500"; dur=@(5,15); weight=3 }
    @{ id=484; cat="bulk"; name="Scroll rows 44500-45000"; dur=@(5,15); weight=3 }
    @{ id=485; cat="bulk"; name="Scroll rows 45000-45500"; dur=@(5,15); weight=3 }
    @{ id=486; cat="bulk"; name="Scroll rows 45500-46000"; dur=@(5,15); weight=3 }
    @{ id=487; cat="bulk"; name="Scroll rows 46000-46500"; dur=@(5,15); weight=3 }
    @{ id=488; cat="bulk"; name="Scroll rows 46500-47000"; dur=@(5,15); weight=3 }
    @{ id=489; cat="bulk"; name="Scroll rows 47000-47500"; dur=@(5,15); weight=3 }
    @{ id=490; cat="bulk"; name="Scroll rows 47500-48000"; dur=@(5,15); weight=3 }
    @{ id=491; cat="bulk"; name="Scroll rows 48000-48500"; dur=@(5,15); weight=3 }
    @{ id=492; cat="bulk"; name="Scroll rows 48500-49000"; dur=@(5,15); weight=3 }
    @{ id=493; cat="bulk"; name="Scroll rows 49000-49500"; dur=@(5,15); weight=3 }
    @{ id=494; cat="bulk"; name="Scroll rows 49500-50000"; dur=@(5,15); weight=3 }
    @{ id=495; cat="bulk"; name="Scroll rows 50000-50500"; dur=@(5,15); weight=3 }
    @{ id=496; cat="bulk"; name="Review data at row 1000"; dur=@(8,20); weight=2 }
    @{ id=497; cat="bulk"; name="Review data at row 2000"; dur=@(8,20); weight=2 }
    @{ id=498; cat="bulk"; name="Review data at row 3000"; dur=@(8,20); weight=2 }
    @{ id=499; cat="bulk"; name="Review data at row 4000"; dur=@(8,20); weight=2 }
    @{ id=500; cat="bulk"; name="Review data at row 5000"; dur=@(8,20); weight=2 }
    @{ id=501; cat="bulk"; name="Review data at row 6000"; dur=@(8,20); weight=2 }
    @{ id=502; cat="bulk"; name="Review data at row 7000"; dur=@(8,20); weight=2 }
    @{ id=503; cat="bulk"; name="Review data at row 8000"; dur=@(8,20); weight=2 }
    @{ id=504; cat="bulk"; name="Review data at row 9000"; dur=@(8,20); weight=2 }
    @{ id=505; cat="bulk"; name="Review data at row 10000"; dur=@(8,20); weight=2 }
    @{ id=506; cat="bulk"; name="Review data at row 11000"; dur=@(8,20); weight=2 }
    @{ id=507; cat="bulk"; name="Review data at row 12000"; dur=@(8,20); weight=2 }
    @{ id=508; cat="bulk"; name="Review data at row 13000"; dur=@(8,20); weight=2 }
    @{ id=509; cat="bulk"; name="Review data at row 14000"; dur=@(8,20); weight=2 }
    @{ id=510; cat="bulk"; name="Review data at row 15000"; dur=@(8,20); weight=2 }
    @{ id=511; cat="bulk"; name="Review data at row 16000"; dur=@(8,20); weight=2 }
    @{ id=512; cat="bulk"; name="Review data at row 17000"; dur=@(8,20); weight=2 }
    @{ id=513; cat="bulk"; name="Review data at row 18000"; dur=@(8,20); weight=2 }
    @{ id=514; cat="bulk"; name="Review data at row 19000"; dur=@(8,20); weight=2 }
    @{ id=515; cat="bulk"; name="Review data at row 20000"; dur=@(8,20); weight=2 }
    @{ id=516; cat="bulk"; name="Review data at row 21000"; dur=@(8,20); weight=2 }
    @{ id=517; cat="bulk"; name="Review data at row 22000"; dur=@(8,20); weight=2 }
    @{ id=518; cat="bulk"; name="Review data at row 23000"; dur=@(8,20); weight=2 }
    @{ id=519; cat="bulk"; name="Review data at row 24000"; dur=@(8,20); weight=2 }
    @{ id=520; cat="bulk"; name="Review data at row 25000"; dur=@(8,20); weight=2 }
    @{ id=521; cat="bulk"; name="Review data at row 26000"; dur=@(8,20); weight=2 }
    @{ id=522; cat="bulk"; name="Review data at row 27000"; dur=@(8,20); weight=2 }
    @{ id=523; cat="bulk"; name="Review data at row 28000"; dur=@(8,20); weight=2 }
    @{ id=524; cat="bulk"; name="Review data at row 29000"; dur=@(8,20); weight=2 }
    @{ id=525; cat="bulk"; name="Review data at row 30000"; dur=@(8,20); weight=2 }
    @{ id=526; cat="bulk"; name="Review data at row 31000"; dur=@(8,20); weight=2 }
    @{ id=527; cat="bulk"; name="Review data at row 32000"; dur=@(8,20); weight=2 }
    @{ id=528; cat="bulk"; name="Review data at row 33000"; dur=@(8,20); weight=2 }
    @{ id=529; cat="bulk"; name="Review data at row 34000"; dur=@(8,20); weight=2 }
    @{ id=530; cat="bulk"; name="Review data at row 35000"; dur=@(8,20); weight=2 }
    @{ id=531; cat="bulk"; name="Review data at row 36000"; dur=@(8,20); weight=2 }
    @{ id=532; cat="bulk"; name="Review data at row 37000"; dur=@(8,20); weight=2 }
    @{ id=533; cat="bulk"; name="Review data at row 38000"; dur=@(8,20); weight=2 }
    @{ id=534; cat="bulk"; name="Review data at row 39000"; dur=@(8,20); weight=2 }
    @{ id=535; cat="bulk"; name="Review data at row 40000"; dur=@(8,20); weight=2 }
    @{ id=536; cat="bulk"; name="Review data at row 41000"; dur=@(8,20); weight=2 }
    @{ id=537; cat="bulk"; name="Review data at row 42000"; dur=@(8,20); weight=2 }
    @{ id=538; cat="bulk"; name="Review data at row 43000"; dur=@(8,20); weight=2 }
    @{ id=539; cat="bulk"; name="Review data at row 44000"; dur=@(8,20); weight=2 }
    @{ id=540; cat="bulk"; name="Review data at row 45000"; dur=@(8,20); weight=2 }
    @{ id=541; cat="bulk"; name="Review data at row 46000"; dur=@(8,20); weight=2 }
    @{ id=542; cat="bulk"; name="Review data at row 47000"; dur=@(8,20); weight=2 }
    @{ id=543; cat="bulk"; name="Review data at row 48000"; dur=@(8,20); weight=2 }
    @{ id=544; cat="bulk"; name="Review data at row 49000"; dur=@(8,20); weight=2 }
    @{ id=545; cat="bulk"; name="Review data at row 50000"; dur=@(8,20); weight=2 }
    @{ id=546; cat="bulk"; name="Review data at row 51000"; dur=@(8,20); weight=2 }
    @{ id=547; cat="bulk"; name="Review data at row 52000"; dur=@(8,20); weight=2 }
    @{ id=548; cat="bulk"; name="Review data at row 53000"; dur=@(8,20); weight=2 }
    @{ id=549; cat="bulk"; name="Review data at row 54000"; dur=@(8,20); weight=2 }
    @{ id=550; cat="bulk"; name="Review data at row 55000"; dur=@(8,20); weight=2 }
    @{ id=551; cat="bulk"; name="Review data at row 56000"; dur=@(8,20); weight=2 }
    @{ id=552; cat="bulk"; name="Review data at row 57000"; dur=@(8,20); weight=2 }
    @{ id=553; cat="bulk"; name="Review data at row 58000"; dur=@(8,20); weight=2 }
    @{ id=554; cat="bulk"; name="Review data at row 59000"; dur=@(8,20); weight=2 }
    @{ id=555; cat="bulk"; name="Review data at row 60000"; dur=@(8,20); weight=2 }
    @{ id=556; cat="bulk"; name="Review data at row 61000"; dur=@(8,20); weight=2 }
    @{ id=557; cat="bulk"; name="Review data at row 62000"; dur=@(8,20); weight=2 }
    @{ id=558; cat="bulk"; name="Review data at row 63000"; dur=@(8,20); weight=2 }
    @{ id=559; cat="bulk"; name="Review data at row 64000"; dur=@(8,20); weight=2 }
    @{ id=560; cat="bulk"; name="Review data at row 65000"; dur=@(8,20); weight=2 }
    @{ id=561; cat="bulk"; name="Review data at row 66000"; dur=@(8,20); weight=2 }
    @{ id=562; cat="bulk"; name="Review data at row 67000"; dur=@(8,20); weight=2 }
    @{ id=563; cat="bulk"; name="Review data at row 68000"; dur=@(8,20); weight=2 }
    @{ id=564; cat="bulk"; name="Review data at row 69000"; dur=@(8,20); weight=2 }
    @{ id=565; cat="bulk"; name="Review data at row 70000"; dur=@(8,20); weight=2 }
    @{ id=566; cat="bulk"; name="Review data at row 71000"; dur=@(8,20); weight=2 }
    @{ id=567; cat="bulk"; name="Review data at row 72000"; dur=@(8,20); weight=2 }
    @{ id=568; cat="bulk"; name="Review data at row 73000"; dur=@(8,20); weight=2 }
    @{ id=569; cat="bulk"; name="Review data at row 74000"; dur=@(8,20); weight=2 }
    @{ id=570; cat="bulk"; name="Review data at row 75000"; dur=@(8,20); weight=2 }
    @{ id=571; cat="bulk"; name="Review data at row 76000"; dur=@(8,20); weight=2 }
    @{ id=572; cat="bulk"; name="Review data at row 77000"; dur=@(8,20); weight=2 }
    @{ id=573; cat="bulk"; name="Review data at row 78000"; dur=@(8,20); weight=2 }
    @{ id=574; cat="bulk"; name="Review data at row 79000"; dur=@(8,20); weight=2 }
    @{ id=575; cat="bulk"; name="Review data at row 80000"; dur=@(8,20); weight=2 }
    @{ id=576; cat="bulk"; name="Review data at row 81000"; dur=@(8,20); weight=2 }
    @{ id=577; cat="bulk"; name="Review data at row 82000"; dur=@(8,20); weight=2 }
    @{ id=578; cat="bulk"; name="Review data at row 83000"; dur=@(8,20); weight=2 }
    @{ id=579; cat="bulk"; name="Review data at row 84000"; dur=@(8,20); weight=2 }
    @{ id=580; cat="bulk"; name="Review data at row 85000"; dur=@(8,20); weight=2 }
    @{ id=581; cat="bulk"; name="Review data at row 86000"; dur=@(8,20); weight=2 }
    @{ id=582; cat="bulk"; name="Review data at row 87000"; dur=@(8,20); weight=2 }
    @{ id=583; cat="bulk"; name="Review data at row 88000"; dur=@(8,20); weight=2 }
    @{ id=584; cat="bulk"; name="Review data at row 89000"; dur=@(8,20); weight=2 }
    @{ id=585; cat="bulk"; name="Review data at row 90000"; dur=@(8,20); weight=2 }
    @{ id=586; cat="bulk"; name="Review data at row 91000"; dur=@(8,20); weight=2 }
    @{ id=587; cat="bulk"; name="Review data at row 92000"; dur=@(8,20); weight=2 }
    @{ id=588; cat="bulk"; name="Review data at row 93000"; dur=@(8,20); weight=2 }
    @{ id=589; cat="bulk"; name="Review data at row 94000"; dur=@(8,20); weight=2 }
    @{ id=590; cat="bulk"; name="Review data at row 95000"; dur=@(8,20); weight=2 }
    @{ id=591; cat="bulk"; name="Review data at row 96000"; dur=@(8,20); weight=2 }
    @{ id=592; cat="bulk"; name="Review data at row 97000"; dur=@(8,20); weight=2 }
    @{ id=593; cat="bulk"; name="Review data at row 98000"; dur=@(8,20); weight=2 }
    @{ id=594; cat="bulk"; name="Review data at row 99000"; dur=@(8,20); weight=2 }
    @{ id=595; cat="bulk"; name="Review data at row 100000"; dur=@(8,20); weight=2 }
    @{ id=596; cat="bulk"; name="Select 100 row range"; dur=@(4,12); weight=2 }
    @{ id=597; cat="bulk"; name="Select 200 row range"; dur=@(4,12); weight=2 }
    @{ id=598; cat="bulk"; name="Select 300 row range"; dur=@(4,12); weight=2 }
    @{ id=599; cat="bulk"; name="Select 400 row range"; dur=@(4,12); weight=2 }
    @{ id=600; cat="bulk"; name="Select 500 row range"; dur=@(4,12); weight=2 }
    @{ id=601; cat="bulk"; name="Select 600 row range"; dur=@(4,12); weight=2 }
    @{ id=602; cat="bulk"; name="Select 700 row range"; dur=@(4,12); weight=2 }
    @{ id=603; cat="bulk"; name="Select 800 row range"; dur=@(4,12); weight=2 }
    @{ id=604; cat="bulk"; name="Select 900 row range"; dur=@(4,12); weight=2 }
    @{ id=605; cat="bulk"; name="Select 1000 row range"; dur=@(4,12); weight=2 }
    @{ id=606; cat="bulk"; name="Select 1100 row range"; dur=@(4,12); weight=2 }
    @{ id=607; cat="bulk"; name="Select 1200 row range"; dur=@(4,12); weight=2 }
    @{ id=608; cat="bulk"; name="Select 1300 row range"; dur=@(4,12); weight=2 }
    @{ id=609; cat="bulk"; name="Select 1400 row range"; dur=@(4,12); weight=2 }
    @{ id=610; cat="bulk"; name="Select 1500 row range"; dur=@(4,12); weight=2 }
    @{ id=611; cat="bulk"; name="Select 1600 row range"; dur=@(4,12); weight=2 }
    @{ id=612; cat="bulk"; name="Select 1700 row range"; dur=@(4,12); weight=2 }
    @{ id=613; cat="bulk"; name="Select 1800 row range"; dur=@(4,12); weight=2 }
    @{ id=614; cat="bulk"; name="Select 1900 row range"; dur=@(4,12); weight=2 }
    @{ id=615; cat="bulk"; name="Select 2000 row range"; dur=@(4,12); weight=2 }
    @{ id=616; cat="bulk"; name="Select 2100 row range"; dur=@(4,12); weight=2 }
    @{ id=617; cat="bulk"; name="Select 2200 row range"; dur=@(4,12); weight=2 }
    @{ id=618; cat="bulk"; name="Select 2300 row range"; dur=@(4,12); weight=2 }
    @{ id=619; cat="bulk"; name="Select 2400 row range"; dur=@(4,12); weight=2 }
    @{ id=620; cat="bulk"; name="Select 2500 row range"; dur=@(4,12); weight=2 }
    @{ id=621; cat="bulk"; name="Select 2600 row range"; dur=@(4,12); weight=2 }
    @{ id=622; cat="bulk"; name="Select 2700 row range"; dur=@(4,12); weight=2 }
    @{ id=623; cat="bulk"; name="Select 2800 row range"; dur=@(4,12); weight=2 }
    @{ id=624; cat="bulk"; name="Select 2900 row range"; dur=@(4,12); weight=2 }
    @{ id=625; cat="bulk"; name="Select 3000 row range"; dur=@(4,12); weight=2 }
    @{ id=626; cat="bulk"; name="Select 3100 row range"; dur=@(4,12); weight=2 }
    @{ id=627; cat="bulk"; name="Select 3200 row range"; dur=@(4,12); weight=2 }
    @{ id=628; cat="bulk"; name="Select 3300 row range"; dur=@(4,12); weight=2 }
    @{ id=629; cat="bulk"; name="Select 3400 row range"; dur=@(4,12); weight=2 }
    @{ id=630; cat="bulk"; name="Select 3500 row range"; dur=@(4,12); weight=2 }
    @{ id=631; cat="bulk"; name="Select 3600 row range"; dur=@(4,12); weight=2 }
    @{ id=632; cat="bulk"; name="Select 3700 row range"; dur=@(4,12); weight=2 }
    @{ id=633; cat="bulk"; name="Select 3800 row range"; dur=@(4,12); weight=2 }
    @{ id=634; cat="bulk"; name="Select 3900 row range"; dur=@(4,12); weight=2 }
    @{ id=635; cat="bulk"; name="Select 4000 row range"; dur=@(4,12); weight=2 }
    @{ id=636; cat="bulk"; name="Select 4100 row range"; dur=@(4,12); weight=2 }
    @{ id=637; cat="bulk"; name="Select 4200 row range"; dur=@(4,12); weight=2 }
    @{ id=638; cat="bulk"; name="Select 4300 row range"; dur=@(4,12); weight=2 }
    @{ id=639; cat="bulk"; name="Select 4400 row range"; dur=@(4,12); weight=2 }
    @{ id=640; cat="bulk"; name="Select 4500 row range"; dur=@(4,12); weight=2 }
    @{ id=641; cat="bulk"; name="Select 4600 row range"; dur=@(4,12); weight=2 }
    @{ id=642; cat="bulk"; name="Select 4700 row range"; dur=@(4,12); weight=2 }
    @{ id=643; cat="bulk"; name="Select 4800 row range"; dur=@(4,12); weight=2 }
    @{ id=644; cat="bulk"; name="Select 4900 row range"; dur=@(4,12); weight=2 }
    @{ id=645; cat="bulk"; name="Select 5000 row range"; dur=@(4,12); weight=2 }
        # Lunch: lighter activity
        return @{ excel=20; bulk=10; perpetua=25; chrome=20; teams=15; human=10 }
    } elseif ($hour -ge 14 -and $hour -lt 17) {
        # Afternoon: optimization work - heavy Excel + bulk operations
        return @{ excel=30; bulk=20; perpetua=25; chrome=10; teams=10; human=5 }
    } else {
        # Default
        return @{ excel=30; bulk=15; perpetua=25; chrome=15; teams=10; human=5 }
    }
}

# ============== DATA POOLS ==============

$script:skuPool = @(
    "NT10234", "NT10567", "NT10891", "JN20345", "JN20678", "JN20912",
    "PR30456", "PR30789", "PR31023", "MG40567", "MG40890", "MG41234",
    "VT50678", "VT50901", "VT51345", "NT11456", "JN21567", "PR31678",
    "MG41789", "VT51890", "NT12567", "JN22678", "PR32789", "MG42890"
)

$script:campaignNames = @(
    "SP_AUTO_TopSellers_Q1", "SP_BRANDED_EXACT_MainKW", "SP_MANUAL_BROAD_Discovery",
    "SP_COMPETITOR_KW_Conquest", "SB_VIDEO_BrandAwareness", "SD_RETARGET_ViewedASIN",
    "SP_AUTO_NewLaunches_Feb", "SP_BRANDED_PHRASE_Secondary", "SP_MANUAL_EXACT_HighIntent",
    "SP_CATEGORY_Targeting_Test", "SB_HEADLINE_Promo_Spring", "SD_AUDIENCE_InMarket",
    "SP_AUTO_LongTail_Explore", "SP_BRANDED_BROAD_Catchall", "SP_MANUAL_NEG_Cleanup"
)

$script:searchQueries = @(
    "amazon ppc acos optimization strategies 2026",
    "perpetua streams bid automation settings guide",
    "sponsored products negative keyword strategy",
    "amazon bulk operations csv format template",
    "branded vs non-branded campaign structure",
    "amazon advertising api rate limits documentation",
    "perpetua goal card custom targeting setup",
    "amazon mcg vs custom goal performance",
    "sponsored brands video creative specs",
    "amazon advertising quarterly report template",
    "acos vs tacos amazon advertising metrics",
    "perpetua dayparting schedule optimization",
    "amazon search term isolation strategy",
    "sponsored display audience targeting guide",
    "amazon ppc budget allocation best practices"
)

$script:websites = @(
    "https://app.perpetua.io/goals",
    "https://app.perpetua.io/streams",
    "https://advertising.amazon.com/cm/campaigns",
    "https://advertising.amazon.com/reports",
    "https://sellercentral.amazon.com/business-reports",
    "https://www.perpetua.io/blog",
    "https://advertising.amazon.com/resources"
)

$script:perpetuaPages = @(
    @{ url="https://app.perpetua.io/am/sp/goals"; name="SP Goals" }
    @{ url="https://app.perpetua.io/am/sp/streams"; name="SP Streams" }
    @{ url="https://app.perpetua.io/am/sb/goals"; name="SB Goals" }
    @{ url="https://app.perpetua.io/am/sd/goals"; name="SD Goals" }
    @{ url="https://app.perpetua.io/am/analytics"; name="Analytics" }
    @{ url="https://app.perpetua.io/am/insights"; name="Insights" }
    @{ url="https://app.perpetua.io/am/reports"; name="Reports" }
    @{ url="https://app.perpetua.io/am/settings"; name="Settings" }
)

$script:perpetuaUI = @{
    sidebar = @{ x=@(30,180); y=@(100,600) }
    sidebarSP = @{ x=100; y=180 }
    sidebarSB = @{ x=100; y=230 }
    sidebarSD = @{ x=100; y=280 }
    sidebarAnalytics = @{ x=100; y=350 }
    header = @{ x=@(200,1800); y=@(20,70) }
    searchBox = @{ x=600; y=45 }
    dateRangePicker = @{ x=1400; y=45 }
    accountDropdown = @{ x=1800; y=45 }
    goalsList = @{ x=@(220,1600); y=@(150,700) }
    goalsTable = @{ x=@(250,1500); y=@(200,650) }
    goalRow = @{ x=@(300,1400); y=@(220,600) }
    goalCard = @{ x=@(300,1200); y=@(150,600) }
    goalMetrics = @{ x=@(800,1100); y=@(200,400) }
    goalTabs = @{ x=@(300,700); y=@(140,170) }
    streamsList = @{ x=@(220,1600); y=@(150,700) }
    streamRow = @{ x=@(250,1500); y=@(180,650) }
    streamFilters = @{ x=@(250,600); y=@(100,140) }
    newGoalBtn = @{ x=1700; y=100 }
    goalNameInput = @{ x=600; y=250 }
    productSearch = @{ x=600; y=350 }
    targetingOptions = @{ x=@(400,900); y=@(400,500) }
    matchTypeCheckboxes = @{ x=@(400,700); y=@(450,520) }
    acosInput = @{ x=700; y=300 }
    budgetInput = @{ x=700; y=360 }
    launchGoalBtn = @{ x=1750; y=50 }
}

# ============== BULK FILE MANAGEMENT ==============

function Initialize-BulkFile {

    try {
        # Try to get existing Excel instance first
        try {
            $script:excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        }
        catch {
            # No Excel running, create new instance
            $script:excel = New-Object -ComObject Excel.Application
            $script:excel.Visible = $true
        }

        $script:excel.DisplayAlerts = $false
        $script:excel.ScreenUpdating = $true

        # Use currently open workbook or create new one
        if ($script:excel.Workbooks.Count -gt 0) {
            $script:workbook = $script:excel.ActiveWorkbook
        }
        else {
            # Try to open the bulk file if it exists
            if (Test-Path $script:bulkFilePath) {
                $script:workbook = $script:excel.Workbooks.Open($script:bulkFilePath, $false, $false)
            }
            else {
                # Create blank workbook
                $script:workbook = $script:excel.Workbooks.Add()
            }
        }

        # Get first worksheet
        $script:worksheet = $script:workbook.Worksheets.Item(1)

        # Analyze file structure

        # Get approximate row count (find last used row)
        $lastRow = $script:worksheet.UsedRange.Rows.Count
        $script:bulkFileInfo.totalRows = $lastRow

        # Get worksheet names
        foreach ($ws in $script:workbook.Worksheets) {
            $script:bulkFileInfo.worksheets += $ws.Name
        }

        # Define useful row ranges for navigation
        $script:bulkFileInfo.rowRanges = @(
            @{ start=1000; end=2000 }
            @{ start=5000; end=6000 }
            @{ start=10000; end=11000 }
            @{ start=25000; end=26000 }
            @{ start=50000; end=51000 }
            @{ start=75000; end=76000 }
            @{ start=100000; end=101000 }
        )

        $script:bulkFileLoaded = $true


        # Move to top of file
        $script:worksheet.Cells.Item(1, 1).Select() | Out-Null

        return $true
    }
    catch {
        Write-Host "[ERROR] Failed to load bulk file: $_" -ForegroundColor Red
        $script:bulkFileLoaded = $false
        return $false
    }
}

function Close-BulkFile {
    if ($script:bulkFileLoaded -and $script:workbook) {
        try {
            $script:workbook.Close($false)  # Don't save changes
            $script:excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:excel) | Out-Null
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        catch {
            Write-Host "[ERROR] Failed to close bulk file: $_" -ForegroundColor Red
        }
    }
}

# ============== MOUSE FUNCTIONS ==============

function Move-MouseSmooth {
    param([int]$targetX, [int]$targetY, [int]$steps = 0)
    $currentPos = [System.Windows.Forms.Cursor]::Position
    $startX = $currentPos.X; $startY = $currentPos.Y

    if ($steps -eq 0) {
        $distance = [Math]::Sqrt([Math]::Pow($targetX - $startX, 2) + [Math]::Pow($targetY - $startY, 2))
        $steps = [Math]::Max(10, [Math]::Min(50, [int]($distance / 20)))
    }

    $wobbleAmount = Get-Random -Minimum 5 -Maximum 20
    $wobbleDirection = if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) { 1 } else { -1 }

    for ($i = 1; $i -le $steps; $i++) {
        $progress = $i / $steps
        $easedProgress = 1 - [Math]::Pow(1 - $progress, 2)
        $wobble = 0
        if ($progress -gt 0.2 -and $progress -lt 0.8) {
            $wobble = [Math]::Sin($progress * [Math]::PI) * $wobbleAmount * $wobbleDirection
        }
        $newX = [int]($startX + ($targetX - $startX) * $easedProgress + $wobble)
        $newY = [int]($startY + ($targetY - $startY) * $easedProgress)
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($newX, $newY)
        $delay = if ($progress -lt 0.2 -or $progress -gt 0.8) { Get-Random -Minimum 8 -Maximum 20 } else { Get-Random -Minimum 3 -Maximum 10 }
        Start-Sleep -Milliseconds $delay
    }
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($targetX, $targetY)
}

function Click-Mouse {
    Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 150)
    [Mouse]::mouse_event([Mouse]::MOUSEEVENTF_LEFTDOWN, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 120)
    [Mouse]::mouse_event([Mouse]::MOUSEEVENTF_LEFTUP, 0, 0, 0, [IntPtr]::Zero)
    Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
}

function Move-AndClick { param([int]$x, [int]$y); Move-MouseSmooth $x $y; Click-Mouse }

function Scroll-MouseWheel {
    param([int]$amount = -3)
    $scrolls = [Math]::Abs($amount)
    $direction = if ($amount -lt 0) { -120 } else { 120 }
    for ($i = 0; $i -lt $scrolls; $i++) {
        [Mouse]::mouse_event([Mouse]::MOUSEEVENTF_WHEEL, 0, 0, $direction, [IntPtr]::Zero)
        Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
    }
}

function Fidget-Mouse {
    $currentPos = [System.Windows.Forms.Cursor]::Position
    $fidgetX = [Math]::Max(10, [Math]::Min($script:screenWidth - 10, $currentPos.X + (Get-Random -Minimum -30 -Maximum 30)))
    $fidgetY = [Math]::Max(10, [Math]::Min($script:screenHeight - 10, $currentPos.Y + (Get-Random -Minimum -20 -Maximum 20)))
    Move-MouseSmooth $fidgetX $fidgetY 8
}

function Get-RandomScreenPosition {
    param([string]$area = "center")
    switch ($area) {
        "excel-cells" { $x = Get-Random -Minimum 100 -Maximum 900; $y = Get-Random -Minimum 150 -Maximum 600 }
        "toolbar" { $x = Get-Random -Minimum 50 -Maximum 800; $y = Get-Random -Minimum 30 -Maximum 120 }
        "address-bar" { $x = Get-Random -Minimum 200 -Maximum 700; $y = Get-Random -Minimum 50 -Maximum 80 }
        "webpage" { $x = Get-Random -Minimum 100 -Maximum 1000; $y = Get-Random -Minimum 200 -Maximum 700 }
        "chat-list" { $x = Get-Random -Minimum 50 -Maximum 280; $y = Get-Random -Minimum 150 -Maximum 600 }
        "chat-area" { $x = Get-Random -Minimum 350 -Maximum 900; $y = Get-Random -Minimum 200 -Maximum 650 }
        "sidebar" { $x = Get-Random -Minimum 20 -Maximum 200; $y = Get-Random -Minimum 100 -Maximum 500 }
        default { $x = Get-Random -Minimum 200 -Maximum 1000; $y = Get-Random -Minimum 200 -Maximum 600 }
    }
    return @{ X = $x; Y = $y }
}

# ============== TYPING FUNCTIONS ==============

function Type-WithTypos {
    param([string]$text, [int]$typoChance = 8)
    foreach ($char in $text.ToCharArray()) {
        if ((Get-Random -Minimum 1 -Maximum 100) -le $typoChance) {
            $wrongChar = [char]((Get-Random -Minimum 97 -Maximum 122))
            [System.Windows.Forms.SendKeys]::SendWait($wrongChar)
            Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
            Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 500)
            [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")
            Start-Sleep -Milliseconds (Get-Random -Minimum 80 -Maximum 200)
        }
        $escaped = $char
        if ($char -match '[\+\^\%\~\(\)\{\}\[\]]') { $escaped = "{$char}" }
        [System.Windows.Forms.SendKeys]::SendWait($escaped)
        $delay = Get-Random -Minimum 50 -Maximum 200
        if ((Get-Random -Minimum 1 -Maximum 100) -le 5) { $delay = Get-Random -Minimum 300 -Maximum 800 }
        Start-Sleep -Milliseconds $delay
    }
}

function Type-Number {
    param([string]$num)
    Type-WithTypos $num 5
}

# ============== TASK SELECTOR (Anti-Pattern) ==============

function Select-NextTask {
    $catWeights = Get-CategoryWeights

    # Build weighted pool excluding recent tasks
    $pool = @()
    foreach ($task in $script:microTasks) {
        # Skip if in last 20 tasks
        if ($script:taskHistory -contains $task.id) { continue }

        # Skip bulk tasks if file not loaded
        if ($task.cat -eq "bulk" -and -not $script:bulkFileLoaded) { continue }

        # Weight by category and task weight
        $catWeight = $catWeights[$task.cat]
        if (-not $catWeight) { $catWeight = 5 }
        $totalWeight = $task.weight * ($catWeight / 10)

        for ($i = 0; $i -lt [Math]::Ceiling($totalWeight); $i++) {
            $pool += $task
        }
    }

    if ($pool.Count -eq 0) {
        # Reset history if pool empty
        $script:taskHistory = @()
        $pool = $script:microTasks | Where-Object { $_.cat -ne "bulk" -or $script:bulkFileLoaded }
    }

    $selected = $pool | Get-Random

    # Update history (keep last 20)
    $script:taskHistory += $selected.id
    if ($script:taskHistory.Count -gt 20) {
        $script:taskHistory = $script:taskHistory[-20..-1]
    }

    return $selected
}

# ============== TASK EXECUTORS ==============

function Execute-BulkTask {
    param($task)

    if (-not $script:bulkFileLoaded) { return }

    try {
        switch -Wildcard ($task.name) {
            "*Jump to row 50000*" {
                $row = 50000
                Write-Host "  [BULK] Jumping to row $row..." -ForegroundColor DarkCyan
                $script:worksheet.Cells.Item($row, 1).Select() | Out-Null
                Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2500)
            }
            "*Jump to row 100000*" {
                $row = 100000
                Write-Host "  [BULK] Jumping to row $row..." -ForegroundColor DarkCyan
                $script:worksheet.Cells.Item($row, 1).Select() | Out-Null
                Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 3000)
            }
            "*Scroll through bulk*" {
                Write-Host "  [BULK] Scrolling through data..." -ForegroundColor DarkCyan
                $range = $script:bulkFileInfo.rowRanges | Get-Random
                $startRow = $range.start
                $script:worksheet.Cells.Item($startRow, 1).Select() | Out-Null
                Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)

                # Scroll down through range
                for ($i = 0; $i -lt 5; $i++) {
                    [System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
                    Start-Sleep -Milliseconds (Get-Random -Minimum 800 -Maximum 2000)
                }
            }
            "*Filter ACOS*" {
                Write-Host "  [BULK] Filtering ACOS > 20%..." -ForegroundColor DarkCyan
                # Click AutoFilter button (Data ribbon location)
                [System.Windows.Forms.SendKeys]::SendWait("%A")  # Alt+A for Data ribbon
                Start-Sleep -Milliseconds 500
                [System.Windows.Forms.SendKeys]::SendWait("T")  # Filter toggle
                Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 3000)
            }
            "*Filter by campaign*" {
                Write-Host "  [BULK] Filtering campaigns..." -ForegroundColor DarkCyan
                [System.Windows.Forms.SendKeys]::SendWait("^f")
                Start-Sleep -Milliseconds 500
                Type-WithTypos ($script:campaignNames | Get-Random).Split("_")[0]
                Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
                [System.Windows.Forms.SendKeys]::SendWait("{ESCAPE}")
            }
            "*Sort by Spend*" {
                Write-Host "  [BULK] Sorting by Spend..." -ForegroundColor DarkCyan
                $script:worksheet.Cells.Item(1, 1).Select() | Out-Null
                [System.Windows.Forms.SendKeys]::SendWait("%A")
                Start-Sleep -Milliseconds 300
                [System.Windows.Forms.SendKeys]::SendWait("SS")  # Sort
                Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
            }
            "*Sort by ACOS*" {
                Write-Host "  [BULK] Sorting by ACOS..." -ForegroundColor DarkCyan
                $script:worksheet.Cells.Item(1, 1).Select() | Out-Null
                [System.Windows.Forms.SendKeys]::SendWait("%A")
                Start-Sleep -Milliseconds 300
                [System.Windows.Forms.SendKeys]::SendWait("SS")
                Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
            }
            "*Clear all filters*" {
                Write-Host "  [BULK] Clearing filters..." -ForegroundColor DarkCyan
                [System.Windows.Forms.SendKeys]::SendWait("%AC")  # Alt+A, C = Clear
                Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
            }
            "*Navigate to Sheet 2*" {
                if ($script:bulkFileInfo.worksheets.Count -gt 1) {
                    Write-Host "  [BULK] Switching worksheet..." -ForegroundColor DarkCyan
                    $script:worksheet = $script:workbook.Worksheets.Item(2)
                    $script:worksheet.Activate()
                    Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2500)
                }
            }
            "*Search specific SKU*" {
                Write-Host "  [BULK] Searching for SKU..." -ForegroundColor DarkCyan
                [System.Windows.Forms.SendKeys]::SendWait("^f")
                Start-Sleep -Milliseconds 500
                Type-WithTypos ($script:skuPool | Get-Random)
                [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
                Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 3000)
                [System.Windows.Forms.SendKeys]::SendWait("{ESCAPE}")
            }
            "*Review data range*" {
                Write-Host "  [BULK] Reviewing data range..." -ForegroundColor DarkCyan
                $range = $script:bulkFileInfo.rowRanges | Get-Random
                $script:worksheet.Cells.Item($range.start, 1).Select() | Out-Null
                Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 5000)

                # Scroll and review
                for ($i = 0; $i -lt 3; $i++) {
                    Scroll-MouseWheel -3
                    Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2500)
                }
            }
            "*Select large range*" {
                Write-Host "  [BULK] Selecting range..." -ForegroundColor DarkCyan
                $range = $script:bulkFileInfo.rowRanges | Get-Random
                $startRow = $range.start
                $endRow = $range.end
                $script:worksheet.Range("A$startRow", "Z$endRow").Select() | Out-Null
                Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2500)
            }
            "*Copy bulk data*" {
                Write-Host "  [BULK] Copying range..." -ForegroundColor DarkCyan
                $range = $script:bulkFileInfo.rowRanges | Get-Random
                $startRow = $range.start
                $endRow = [Math]::Min($startRow + 100, $range.end)
                $script:worksheet.Range("A$startRow", "E$endRow").Select() | Out-Null
                Start-Sleep -Milliseconds 500
                [System.Windows.Forms.SendKeys]::SendWait("^c")
                Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
            }
            "*end of data*" {
                Write-Host "  [BULK] Jumping to end..." -ForegroundColor DarkCyan
                [System.Windows.Forms.SendKeys]::SendWait("^{END}")
                Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2500)
            }
            "*Return to top*" {
                Write-Host "  [BULK] Returning to top..." -ForegroundColor DarkCyan
                [System.Windows.Forms.SendKeys]::SendWait("^{HOME}")
                Start-Sleep -Milliseconds (Get-Random -Minimum 800 -Maximum 1500)
            }
        }
    }
    catch {
        Write-Host "  [ERROR] Bulk task failed: $_" -ForegroundColor Red
    }
}

function Execute-ExcelTask {
    param($task)

    $pos = Get-RandomScreenPosition "excel-cells"

    switch -Wildcard ($task.name) {
        "*SKU*" {
            Move-AndClick $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 500)
            Type-WithTypos ($script:skuPool | Get-Random)
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
        }
        "*campaign name*" {
            Move-AndClick $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 500)
            Type-WithTypos ($script:campaignNames | Get-Random)
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        }
        "*ACOS*" {
            Move-AndClick $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 400)
            $acos = [math]::Round((Get-Random -Minimum 800 -Maximum 6500) / 100, 2)
            Type-Number "$acos%"
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
        }
        "*spend*" {
            Move-AndClick $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 400)
            $spend = [math]::Round((Get-Random -Minimum 500 -Maximum 50000) / 100, 2)
            Type-Number "`$$spend"
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
        }
        "*bid*" {
            Move-AndClick $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 400)
            $bid = [math]::Round((Get-Random -Minimum 15 -Maximum 350) / 100, 2)
            Type-Number "`$$bid"
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        }
        "*formula*" {
            Move-AndClick $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 600)
            $formulas = @("=SUM(B2:B50)", "=AVERAGE(C2:C100)", "=B2/C2", "=D2*0.15", "=IF(E2>30,""HIGH"",""OK"")")
            Type-WithTypos ($formulas | Get-Random)
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        }
        "*Scroll*" {
            Move-MouseSmooth $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 600)
            $dir = if ($task.name -match "up") { 3 } else { -3 }
            Scroll-MouseWheel $dir
        }
        "*Click*cell*" {
            Move-AndClick $pos.X $pos.Y
        }
        "*Tab*" {
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
        }
        "*Enter*" {
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        }
        "*Save*" {
            Fidget-Mouse
            Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 400)
            [System.Windows.Forms.SendKeys]::SendWait("^s")
        }
        "*Find*" {
            [System.Windows.Forms.SendKeys]::SendWait("^f")
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
            Type-WithTypos ($script:skuPool | Get-Random).Substring(0,4)
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
            [System.Windows.Forms.SendKeys]::SendWait("{ESCAPE}")
        }
        "*Copy*" {
            Move-AndClick $pos.X $pos.Y
            [System.Windows.Forms.SendKeys]::SendWait("^c")
        }
        "*Paste*" {
            Move-AndClick $pos.X $pos.Y
            [System.Windows.Forms.SendKeys]::SendWait("^v")
        }
        "*Undo*" {
            [System.Windows.Forms.SendKeys]::SendWait("^z")
        }
        "*worksheet*" {
            $tabX = Get-Random -Minimum 100 -Maximum 400
            Move-AndClick $tabX ($script:screenHeight - 60)
        }
        default {
            Move-AndClick $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 800)
            if ((Get-Random -Minimum 1 -Maximum 100) -le 50) {
                $num = Get-Random -Minimum 100 -Maximum 9999
                Type-Number $num.ToString()
                [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
            }
        }
    }
}

function Execute-ChromeTask {
    param($task)

    switch -Wildcard ($task.name) {
        "*search*" {
            $pos = Get-RandomScreenPosition "address-bar"
            Move-MouseSmooth $pos.X $pos.Y
            [System.Windows.Forms.SendKeys]::SendWait("^a")
            Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
            Type-WithTypos ($script:searchQueries | Get-Random)
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 5000)
        }
        "*Load*" {
            $pos = Get-RandomScreenPosition "address-bar"
            Move-MouseSmooth $pos.X $pos.Y
            [System.Windows.Forms.SendKeys]::SendWait("^a")
            Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
            Type-WithTypos ($script:websites | Get-Random)
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 5000)
        }
        "*Scroll*" {
            $pos = Get-RandomScreenPosition "webpage"
            Move-MouseSmooth $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 800)
            $scrollAmt = Get-Random -Minimum -5 -Maximum -2
            Scroll-MouseWheel $scrollAmt
        }
        "*Click*" {
            # Just mouse pan, no click
            $pos = Get-RandomScreenPosition "webpage"
            Move-MouseSmooth $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1200)
        }
        "*new tab*" {
            # Skip - don't open tabs
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 800)
        }
        "*Close tab*" {
            # Skip - don't close tabs
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 800)
        }
        "*Switch tab*" {
            [System.Windows.Forms.SendKeys]::SendWait("^{TAB}")
        }
        "*back*" {
            # Just scroll up instead of going back
            Scroll-MouseWheel 3
        }
        "*Refresh*" {
            # Skip refresh - just mouse pan
            $pos = Get-RandomScreenPosition "webpage"
            Move-MouseSmooth $pos.X $pos.Y
        }
        default {
            # Default: just scroll and mouse pan
            $pos = Get-RandomScreenPosition "webpage"
            Move-MouseSmooth $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
            Scroll-MouseWheel (Get-Random -Minimum -3 -Maximum 3)
        }
    }
}

function Execute-PerpetuaTask {
    param($task)

    $ui = $script:perpetuaUI

    switch -Wildcard ($task.name) {
        "*Navigate SP Goals*" {
            # Just mouse pan to sidebar area
            Move-MouseSmooth $ui.sidebarSP.x $ui.sidebarSP.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*Navigate SB Goals*" {
            Move-MouseSmooth $ui.sidebarSB.x $ui.sidebarSB.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*Navigate SD Goals*" {
            Move-MouseSmooth $ui.sidebarSD.x $ui.sidebarSD.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*Click goal row*" {
            # Just hover over row area, no click
            $x = Get-Random -Minimum $ui.goalRow.x[0] -Maximum $ui.goalRow.x[1]
            $y = Get-Random -Minimum $ui.goalRow.y[0] -Maximum $ui.goalRow.y[1]
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*Review goal*metric*" {
            $x = Get-Random -Minimum $ui.goalMetrics.x[0] -Maximum $ui.goalMetrics.x[1]
            $y = Get-Random -Minimum $ui.goalMetrics.y[0] -Maximum $ui.goalMetrics.y[1]
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 5000)
            Fidget-Mouse
        }
        "*Scroll goals list*" {
            $x = Get-Random -Minimum $ui.goalsList.x[0] -Maximum $ui.goalsList.x[1]
            $y = Get-Random -Minimum $ui.goalsList.y[0] -Maximum $ui.goalsList.y[1]
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 600)
            $dir = if ($task.name -match "up") { 3 } else { -3 }
            Scroll-MouseWheel $dir
        }
        "*Click goal tabs*" {
            # Just hover over tabs area
            $x = Get-Random -Minimum $ui.goalTabs.x[0] -Maximum $ui.goalTabs.x[1]
            $y = Get-Random -Minimum $ui.goalTabs.y[0] -Maximum $ui.goalTabs.y[1]
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 800 -Maximum 1500)
        }
        "*Search goal*" {
            # Just hover over search area, no typing
            Move-MouseSmooth $ui.searchBox.x $ui.searchBox.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*Filter goals*" {
            # Just mouse pan in filter area
            $x = Get-Random -Minimum 300 -Maximum 500
            Move-MouseSmooth $x 120
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*Sort by*column*" {
            # Hover over column header
            $x = if ($task.name -match "ACOS") { 900 } else { 800 }
            Move-MouseSmooth $x 175
            Start-Sleep -Milliseconds (Get-Random -Minimum 800 -Maximum 1500)
        }
        "*Change date range*" {
            # Just hover over date picker
            Move-MouseSmooth $ui.dateRangePicker.x $ui.dateRangePicker.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*New Goal button*" {
            # Just hover, don't click
            Move-MouseSmooth $ui.newGoalBtn.x $ui.newGoalBtn.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*Close goal detail*" {
            [System.Windows.Forms.SendKeys]::SendWait("{ESCAPE}")
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
        }
        "*Navigate SP Streams*" {
            $x = Get-Random -Minimum $ui.sidebar.x[0] -Maximum $ui.sidebar.x[1]
            Move-MouseSmooth $x 320
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
        }
        "*Click stream row*" {
            $x = Get-Random -Minimum $ui.streamRow.x[0] -Maximum $ui.streamRow.x[1]
            $y = Get-Random -Minimum $ui.streamRow.y[0] -Maximum $ui.streamRow.y[1]
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2500)
        }
        "*Scroll streams list*" {
            $x = Get-Random -Minimum $ui.streamsList.x[0] -Maximum $ui.streamsList.x[1]
            $y = Get-Random -Minimum $ui.streamsList.y[0] -Maximum $ui.streamsList.y[1]
            Move-MouseSmooth $x $y
            Scroll-MouseWheel (Get-Random -Minimum -4 -Maximum -2)
        }
        "*View stream bid*" {
            $x = Get-Random -Minimum 600 -Maximum 900
            $y = Get-Random -Minimum 250 -Maximum 400
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 5000)
        }
        "*Filter streams*" {
            $x = Get-Random -Minimum $ui.streamFilters.x[0] -Maximum $ui.streamFilters.x[1]
            $y = Get-Random -Minimum $ui.streamFilters.y[0] -Maximum $ui.streamFilters.y[1]
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*stream*trend*" {
            $x = Get-Random -Minimum 700 -Maximum 1200
            $y = Get-Random -Minimum 300 -Maximum 450
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 4000)
            Fidget-Mouse
        }
        "*Expand stream*" {
            $x = Get-Random -Minimum 250 -Maximum 300
            $y = Get-Random -Minimum 220 -Maximum 500
            Move-MouseSmooth $x $y
        }
        "*Collapse stream*" {
            $x = Get-Random -Minimum 250 -Maximum 300
            $y = Get-Random -Minimum 220 -Maximum 400
            Move-MouseSmooth $x $y
        }
        "*Sort streams*" {
            $x = Get-Random -Minimum 700 -Maximum 900
            Move-MouseSmooth $x 175
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*stream automation*" {
            $x = Get-Random -Minimum 1100 -Maximum 1300
            $y = Get-Random -Minimum 250 -Maximum 450
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 3000)
        }
        "*Navigate Analytics*" {
            Move-MouseSmooth $ui.sidebarAnalytics.x $ui.sidebarAnalytics.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
        }
        "*sidebar nav*" {
            $x = Get-Random -Minimum $ui.sidebar.x[0] -Maximum $ui.sidebar.x[1]
            $y = Get-Random -Minimum $ui.sidebar.y[0] -Maximum $ui.sidebar.y[1]
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 3000)
        }
        "*Hover sidebar*" {
            $x = Get-Random -Minimum 20 -Maximum 60
            $y = Get-Random -Minimum 150 -Maximum 500
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 800 -Maximum 2000)
        }
        "*account dropdown*" {
            Move-MouseSmooth $ui.accountDropdown.x $ui.accountDropdown.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
            Move-MouseSmooth 800 400
        }
        "*notifications bell*" {
            Move-MouseSmooth ($ui.accountDropdown.x - 80) $ui.accountDropdown.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2500)
            Move-MouseSmooth 800 400
        }
        "*Refresh current*" {
            [System.Windows.Forms.SendKeys]::SendWait("{F5}")
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
        }
        "*breadcrumb*" {
            $x = Get-Random -Minimum 220 -Maximum 400
            Move-MouseSmooth $x 80
            Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 3000)
        }
        "*Scroll page randomly*" {
            $x = Get-Random -Minimum 400 -Maximum 1200
            $y = Get-Random -Minimum 300 -Maximum 600
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 800)
            $dir = if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) { -3 } else { 3 }
            Scroll-MouseWheel $dir
        }
        "*idle on metrics*" {
            $x = Get-Random -Minimum 600 -Maximum 1100
            $y = Get-Random -Minimum 200 -Maximum 500
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 3000 -Maximum 8000)
            if ((Get-Random -Minimum 1 -Maximum 100) -le 40) { Fidget-Mouse }
        }
        default {
            $x = Get-Random -Minimum 300 -Maximum 1200
            $y = Get-Random -Minimum 200 -Maximum 600
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1500)
        }
    }
}

function Execute-TeamsTask {
    param($task)

    switch -Wildcard ($task.name) {
        "*chat*" {
            $pos = Get-RandomScreenPosition "chat-list"
            Move-AndClick $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
            $chatPos = Get-RandomScreenPosition "chat-area"
            Move-MouseSmooth $chatPos.X $chatPos.Y
        }
        "*Read*" {
            $pos = Get-RandomScreenPosition "chat-area"
            Move-MouseSmooth $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 5000)
        }
        "*Scroll*" {
            $pos = if ($task.name -match "channel") { Get-RandomScreenPosition "sidebar" } else { Get-RandomScreenPosition "chat-area" }
            Move-MouseSmooth $pos.X $pos.Y
            Scroll-MouseWheel (Get-Random -Minimum -3 -Maximum 3)
        }
        "*activity*" {
            $pos = Get-RandomScreenPosition "sidebar"
            Move-AndClick $pos.X 100
        }
        "*React*" {
            $pos = Get-RandomScreenPosition "chat-area"
            Move-MouseSmooth $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
        }
        default {
            $pos = Get-RandomScreenPosition "chat-list"
            Move-AndClick $pos.X $pos.Y
        }
    }
}

function Execute-HumanTask {
    param($task)

    switch -Wildcard ($task.name) {
        "*Pause*think*" {
            if ((Get-Random -Minimum 1 -Maximum 100) -le 30) { Fidget-Mouse }
        }
        "*fidget*" {
            Fidget-Mouse
        }
        "*Hesitation*" {
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 2000)
            if ((Get-Random -Minimum 1 -Maximum 100) -le 40) { Fidget-Mouse }
        }
        "*Re-read*" {
            $pos = Get-RandomScreenPosition "center"
            Move-MouseSmooth $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 3000)
        }
        "*break*" {
            $breakTime = Get-Random -Minimum 15 -Maximum 45
            for ($i = 0; $i -lt $breakTime; $i += 5) {
                Start-Sleep -Seconds 5
                if ((Get-Random -Minimum 1 -Maximum 100) -le 20) { Fidget-Mouse }
            }
        }
    }
}

# ============== APP SWITCHING ==============

function Switch-ToApp {
    param([string]$targetApp)

    if ($targetApp -eq $script:lastApp) { return }
    if ($targetApp -eq "bulk") { $targetApp = "excel" }

    # Use WScript to activate window reliably (no Alt+Tab!)
    $wshell = New-Object -ComObject wscript.shell

    switch ($targetApp) {
        "excel" {
            if ($script:excel) {
                $wshell.AppActivate($script:excel.Caption) | Out-Null
            } else {
                $wshell.AppActivate("Excel") | Out-Null
            }
        }
        "perpetua" {
            $wshell.AppActivate("Perpetua") | Out-Null
            if (-not $?) { $wshell.AppActivate("Chrome") | Out-Null }
        }
        "chrome" {
            $wshell.AppActivate("Chrome") | Out-Null
        }
        "teams" {
            $wshell.AppActivate("Teams") | Out-Null
        }
    }

    Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1500)
    $script:lastApp = $targetApp

    # 5 combined movements after window switch
    for ($i = 0; $i -lt 5; $i++) {
        $pos = Get-RandomScreenPosition "center"
        Move-MouseSmooth $pos.X $pos.Y (Get-Random -Minimum 5 -Maximum 15)
        Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 500)
    }
}

# ============== MAIN TRACKER ==============

function Start-Tracking {
    $totalActions = 0
    $script:sessionStart = Get-Date


    # Initialize bulk file
    $bulkLoaded = Initialize-BulkFile

    if ($bulkLoaded) {
    } else {
        Write-Host "[WARNING] Bulk file not loaded. Running without bulk tasks." -ForegroundColor Yellow
    }

    Start-Sleep -Seconds 3

    $pos = Get-RandomScreenPosition "excel-cells"
    Move-MouseSmooth $pos.X $pos.Y
    $script:lastApp = "excel"

    try {
        while ($true) {
            # Execute 5 tasks in immediate succession (like human working)
            for ($burst = 0; $burst -lt 5; $burst++) {
                $task = Select-NextTask

                # Switch app only at start of burst
                $taskCat = if ($task.cat -eq "bulk") { "excel" } else { $task.cat }
                if ($taskCat -ne $script:lastApp -and $taskCat -ne "human" -and $burst -eq 0) {
                    Switch-ToApp $taskCat
                }

                # Execute task immediately (no gap during burst)
                switch ($task.cat) {
                    "excel" { Execute-ExcelTask $task }
                    "bulk" { Execute-BulkTask $task }
                    "perpetua" { Execute-PerpetuaTask $task }
                    "chrome" { Execute-ChromeTask $task }
                    "teams" { Execute-TeamsTask $task }
                    "human" { Execute-HumanTask $task }
                }

                $totalActions++

                # NO gap during burst - tasks execute back-to-back
                # Just a tiny natural delay (human can't be instant)
                Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 600)
            }

            # Now pause after completing all 5 tasks (1-15 sec)
            $pause = Get-Random -Minimum 1 -Maximum 15
            Start-Sleep -Seconds $pause

            # Occasional fidget during pause
            if ((Get-Random -Maximum 100) -le 40) { Fidget-Mouse }
        }
    }
    finally {
        $runTime = [math]::Round(((Get-Date) - $script:sessionStart).TotalMinutes, 1)
        Write-Host "`n[STOPPED] Runtime: $runTime min | Actions: $totalActions" -ForegroundColor Yellow

        # Close bulk file
        Close-BulkFile
    }
}

# ============== CONTROL INTERFACE ==============

Clear-Host
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Campaign Processor" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Time-of-day awareness | Human behaviors"
Write-Host ""
Write-Host "Commands:"
Write-Host "  go   - Start tracking (processes data)"
Write-Host "  exit - Close"
Write-Host ""
Write-Host "----------------------------------------"
Write-Host ""

while ($true) {
    $cmd = Read-Host "tracker"
    switch ($cmd.ToLower().Trim()) {
        "go" { Start-Tracking; Write-Host "" }
        "exit" {
            Close-BulkFile
            exit
        }
        default { if ($cmd -ne "") { Write-Host "[ERROR] Unknown: $cmd (use: go | exit)" -ForegroundColor Red } }
    }
}
