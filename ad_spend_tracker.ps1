# Amazon Advertising Spend Tracker v3.1
# 120+ micro-tasks with precise Perpetua UI interactions

Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Mouse {
    [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, int dwData, IntPtr dwExtraInfo);
    public const uint MOUSEEVENTF_LEFTDOWN = 0x02;
    public const uint MOUSEEVENTF_LEFTUP = 0x04;
    public const uint MOUSEEVENTF_WHEEL = 0x0800;
}
"@

# ============== CONFIGURATION ==============

$script:screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
$script:screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height

# Task history for anti-pattern (never repeat last 20)
$script:taskHistory = @()
$script:lastApp = ""
$script:sessionStart = Get-Date

# ============== 100+ MICRO-TASKS ==============

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
        # Morning: more reporting/review - heavy Perpetua
        return @{ excel=30; perpetua=35; chrome=20; teams=10; human=5 }
    } elseif ($hour -ge 12 -and $hour -lt 14) {
        # Lunch: lighter activity
        return @{ excel=25; perpetua=25; chrome=20; teams=20; human=10 }
    } elseif ($hour -ge 14 -and $hour -lt 17) {
        # Afternoon: optimization work - heavy Excel
        return @{ excel=40; perpetua=30; chrome=15; teams=10; human=5 }
    } else {
        # Default
        return @{ excel=35; perpetua=30; chrome=18; teams=12; human=5 }
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

# Perpetua-specific pages for realistic navigation
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

# Perpetua UI element positions (approximate, based on 1920x1080)
$script:perpetuaUI = @{
    # Sidebar navigation (left side)
    sidebar = @{ x=@(30,180); y=@(100,600) }
    sidebarSP = @{ x=100; y=180 }
    sidebarSB = @{ x=100; y=230 }
    sidebarSD = @{ x=100; y=280 }
    sidebarAnalytics = @{ x=100; y=350 }

    # Top navigation/header
    header = @{ x=@(200,1800); y=@(20,70) }
    searchBox = @{ x=600; y=45 }
    dateRangePicker = @{ x=1400; y=45 }
    accountDropdown = @{ x=1800; y=45 }

    # Goals list area
    goalsList = @{ x=@(220,1600); y=@(150,700) }
    goalsTable = @{ x=@(250,1500); y=@(200,650) }
    goalRow = @{ x=@(300,1400); y=@(220,600) }

    # Goal card/detail view
    goalCard = @{ x=@(300,1200); y=@(150,600) }
    goalMetrics = @{ x=@(800,1100); y=@(200,400) }
    goalTabs = @{ x=@(300,700); y=@(140,170) }

    # Streams page
    streamsList = @{ x=@(220,1600); y=@(150,700) }
    streamRow = @{ x=@(250,1500); y=@(180,650) }
    streamFilters = @{ x=@(250,600); y=@(100,140) }

    # New Goal wizard
    newGoalBtn = @{ x=1700; y=100 }
    goalNameInput = @{ x=600; y=250 }
    productSearch = @{ x=600; y=350 }
    targetingOptions = @{ x=@(400,900); y=@(400,500) }
    matchTypeCheckboxes = @{ x=@(400,700); y=@(450,520) }
    acosInput = @{ x=700; y=300 }
    budgetInput = @{ x=700; y=360 }
    launchGoalBtn = @{ x=1750; y=50 }
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
        $pool = $script:microTasks
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
            # Generic cell interaction
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
            Move-AndClick $pos.X $pos.Y
            [System.Windows.Forms.SendKeys]::SendWait("^a")
            Start-Sleep -Milliseconds (Get-Random -Minimum 100 -Maximum 300)
            Type-WithTypos ($script:searchQueries | Get-Random)
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 5000)
        }
        "*Load*" {
            $pos = Get-RandomScreenPosition "address-bar"
            Move-AndClick $pos.X $pos.Y
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
            $pos = Get-RandomScreenPosition "webpage"
            Move-AndClick $pos.X $pos.Y
        }
        "*new tab*" {
            [System.Windows.Forms.SendKeys]::SendWait("^t")
        }
        "*Close tab*" {
            [System.Windows.Forms.SendKeys]::SendWait("^w")
        }
        "*Switch tab*" {
            [System.Windows.Forms.SendKeys]::SendWait("^{TAB}")
        }
        "*back*" {
            if ((Get-Random -Minimum 0 -Maximum 2) -eq 0) {
                Move-AndClick 45 65
            } else {
                [System.Windows.Forms.SendKeys]::SendWait("%{LEFT}")
            }
        }
        "*Refresh*" {
            [System.Windows.Forms.SendKeys]::SendWait("{F5}")
        }
        default {
            $pos = Get-RandomScreenPosition "webpage"
            Move-AndClick $pos.X $pos.Y
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1500)
        }
    }
}

function Execute-PerpetuaTask {
    param($task)

    # Get Perpetua UI positions
    $ui = $script:perpetuaUI

    switch -Wildcard ($task.name) {
        # === GOALS PAGE INTERACTIONS ===
        "*Navigate SP Goals*" {
            # Click sidebar SP item
            Move-AndClick $ui.sidebarSP.x $ui.sidebarSP.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
        }
        "*Navigate SB Goals*" {
            Move-AndClick $ui.sidebarSB.x $ui.sidebarSB.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
        }
        "*Navigate SD Goals*" {
            Move-AndClick $ui.sidebarSD.x $ui.sidebarSD.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
        }
        "*Click goal row*" {
            $x = Get-Random -Minimum $ui.goalRow.x[0] -Maximum $ui.goalRow.x[1]
            $y = Get-Random -Minimum $ui.goalRow.y[0] -Maximum $ui.goalRow.y[1]
            Move-AndClick $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2500)
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
            $x = Get-Random -Minimum $ui.goalTabs.x[0] -Maximum $ui.goalTabs.x[1]
            $y = Get-Random -Minimum $ui.goalTabs.y[0] -Maximum $ui.goalTabs.y[1]
            Move-AndClick $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 3000)
        }
        "*Search goal*" {
            Move-AndClick $ui.searchBox.x $ui.searchBox.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 600)
            [System.Windows.Forms.SendKeys]::SendWait("^a")
            Start-Sleep -Milliseconds 200
            $goalName = $script:campaignNames | Get-Random
            $searchTerm = $goalName.Split("_")[0..1] -join "_"
            Type-WithTypos $searchTerm
            Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 3000)
        }
        "*Filter goals*" {
            # Click filter dropdown area
            $x = Get-Random -Minimum 300 -Maximum 500
            Move-AndClick $x 120
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
            # Click option
            $y = if ($task.name -match "Enabled") { 160 } else { 190 }
            Move-AndClick ($x + 20) $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*Sort by*column*" {
            # Click column header
            $x = if ($task.name -match "ACOS") { 900 } else { 800 }
            Move-AndClick $x 175
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*Change date range*" {
            Move-AndClick $ui.dateRangePicker.x $ui.dateRangePicker.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
            # Click a preset option
            $y = if ($task.name -match "7d") { 200 } else { 230 }
            Move-AndClick ($ui.dateRangePicker.x - 50) $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 3000)
        }
        "*New Goal button*" {
            Move-AndClick $ui.newGoalBtn.x $ui.newGoalBtn.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
            # Press Escape to close without creating
            [System.Windows.Forms.SendKeys]::SendWait("{ESCAPE}")
            Start-Sleep -Milliseconds 500
        }
        "*Close goal detail*" {
            [System.Windows.Forms.SendKeys]::SendWait("{ESCAPE}")
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
        }

        # === STREAMS PAGE INTERACTIONS ===
        "*Navigate SP Streams*" {
            # Click Streams in sidebar or tab
            $x = Get-Random -Minimum $ui.sidebar.x[0] -Maximum $ui.sidebar.x[1]
            Move-AndClick $x 320
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
        }
        "*Click stream row*" {
            $x = Get-Random -Minimum $ui.streamRow.x[0] -Maximum $ui.streamRow.x[1]
            $y = Get-Random -Minimum $ui.streamRow.y[0] -Maximum $ui.streamRow.y[1]
            Move-AndClick $x $y
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
            Move-AndClick $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*stream*trend*" {
            # Hover over graph area
            $x = Get-Random -Minimum 700 -Maximum 1200
            $y = Get-Random -Minimum 300 -Maximum 450
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 4000)
            Fidget-Mouse
        }
        "*Expand stream*" {
            $x = Get-Random -Minimum 250 -Maximum 300
            $y = Get-Random -Minimum 220 -Maximum 500
            Move-AndClick $x $y
        }
        "*Collapse stream*" {
            $x = Get-Random -Minimum 250 -Maximum 300
            $y = Get-Random -Minimum 220 -Maximum 400
            Move-AndClick $x $y
        }
        "*Sort streams*" {
            $x = Get-Random -Minimum 700 -Maximum 900
            Move-AndClick $x 175
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2000)
        }
        "*stream automation*" {
            # Click automation toggle area
            $x = Get-Random -Minimum 1100 -Maximum 1300
            $y = Get-Random -Minimum 250 -Maximum 450
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 3000)
        }
        "*Navigate Analytics*" {
            Move-AndClick $ui.sidebarAnalytics.x $ui.sidebarAnalytics.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
        }

        # === GENERAL PERPETUA INTERACTIONS ===
        "*sidebar nav*" {
            $x = Get-Random -Minimum $ui.sidebar.x[0] -Maximum $ui.sidebar.x[1]
            $y = Get-Random -Minimum $ui.sidebar.y[0] -Maximum $ui.sidebar.y[1]
            Move-AndClick $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1500 -Maximum 3000)
        }
        "*Hover sidebar*" {
            $x = Get-Random -Minimum 20 -Maximum 60
            $y = Get-Random -Minimum 150 -Maximum 500
            Move-MouseSmooth $x $y
            Start-Sleep -Milliseconds (Get-Random -Minimum 800 -Maximum 2000)
        }
        "*account dropdown*" {
            Move-AndClick $ui.accountDropdown.x $ui.accountDropdown.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1000)
            # Click elsewhere to close
            Move-AndClick 800 400
        }
        "*notifications bell*" {
            Move-AndClick ($ui.accountDropdown.x - 80) $ui.accountDropdown.y
            Start-Sleep -Milliseconds (Get-Random -Minimum 1000 -Maximum 2500)
            # Click elsewhere to close
            Move-AndClick 800 400
        }
        "*Refresh current*" {
            [System.Windows.Forms.SendKeys]::SendWait("{F5}")
            Start-Sleep -Milliseconds (Get-Random -Minimum 2000 -Maximum 4000)
        }
        "*breadcrumb*" {
            $x = Get-Random -Minimum 220 -Maximum 400
            Move-AndClick $x 80
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
            # Just hover and look
            Start-Sleep -Milliseconds (Get-Random -Minimum 3000 -Maximum 8000)
            if ((Get-Random -Minimum 1 -Maximum 100) -le 40) { Fidget-Mouse }
        }
        default {
            # Generic Perpetua interaction
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
            # Hover for reaction
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
            # Just wait, maybe small fidget
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
            # Longer pause with occasional fidget
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

    # Move to taskbar area
    $taskbarY = $script:screenHeight - 30
    $taskbarX = Get-Random -Minimum 200 -Maximum 800
    Move-MouseSmooth $taskbarX $taskbarY
    Start-Sleep -Milliseconds (Get-Random -Minimum 200 -Maximum 500)

    # Alt+Tab cycles
    $tabCount = Get-Random -Minimum 1 -Maximum 3
    for ($i = 0; $i -lt $tabCount; $i++) {
        [System.Windows.Forms.SendKeys]::SendWait("%{TAB}")
        Start-Sleep -Milliseconds (Get-Random -Minimum 300 -Maximum 600)
    }

    Start-Sleep -Milliseconds (Get-Random -Minimum 500 -Maximum 1500)
    $script:lastApp = $targetApp

    # Move to app area
    $pos = Get-RandomScreenPosition "center"
    Move-MouseSmooth $pos.X $pos.Y
}

# ============== MAIN TRACKER ==============

function Start-Tracking {
    $totalActions = 0
    $script:sessionStart = Get-Date

    Write-Host "[TRACKING] Started at $(Get-Date -Format 'HH:mm:ss')" -ForegroundColor Green
    Write-Host "[INFO] v3.1 - 130 micro-tasks, Perpetua UI, anti-pattern" -ForegroundColor Cyan
    Write-Host "[INFO] Click Excel window now. Ctrl+C to stop." -ForegroundColor Yellow
    Start-Sleep -Seconds 3

    $pos = Get-RandomScreenPosition "excel-cells"
    Move-MouseSmooth $pos.X $pos.Y
    $script:lastApp = "excel"

    try {
        while ($true) {
            # Select next task (anti-pattern)
            $task = Select-NextTask

            # Maybe switch app
            if ($task.cat -ne $script:lastApp -and $task.cat -ne "human") {
                Switch-ToApp $task.cat
            }

            # Execute task
            switch ($task.cat) {
                "excel" { Execute-ExcelTask $task }
                "perpetua" { Execute-PerpetuaTask $task }
                "chrome" { Execute-ChromeTask $task }
                "teams" { Execute-TeamsTask $task }
                "human" { Execute-HumanTask $task }
            }

            $totalActions++

            # Task duration (from task spec)
            $minDur = $task.dur[0]
            $maxDur = $task.dur[1]
            $taskDuration = Get-Random -Minimum $minDur -Maximum $maxDur

            # Execute duration with fidgets
            $elapsed = 0
            while ($elapsed -lt $taskDuration) {
                $sleepTime = [Math]::Min(5, $taskDuration - $elapsed)
                Start-Sleep -Seconds $sleepTime
                $elapsed += $sleepTime
                if ((Get-Random -Minimum 1 -Maximum 100) -le 15) { Fidget-Mouse }
            }

            # Gap between tasks (5-29 sec)
            $gap = Get-Random -Minimum 5 -Maximum 29
            $gapElapsed = 0
            while ($gapElapsed -lt $gap) {
                Start-Sleep -Seconds 5
                $gapElapsed += 5
                if ((Get-Random -Minimum 1 -Maximum 100) -le 20) { Fidget-Mouse }
            }
        }
    }
    finally {
        $runTime = [math]::Round(((Get-Date) - $script:sessionStart).TotalMinutes, 1)
        Write-Host "`n[STOPPED] Runtime: $runTime min | Actions: $totalActions" -ForegroundColor Yellow
    }
}

# ============== CONTROL INTERFACE ==============

Clear-Host
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Amazon Advertising Spend Tracker v3.1" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "130 micro-tasks | Perpetua UI precision | Anti-pattern"
Write-Host "Time-of-day awareness | Human behaviors"
Write-Host ""
Write-Host "Commands:"
Write-Host "  go   - Start tracking"
Write-Host "  exit - Close"
Write-Host ""
Write-Host "Setup: Open Excel, Chrome, Teams first"
Write-Host "----------------------------------------"
Write-Host ""

while ($true) {
    $cmd = Read-Host "tracker"
    switch ($cmd.ToLower().Trim()) {
        "go" { Start-Tracking; Write-Host "" }
        "exit" { Write-Host "[INFO] Closing..." -ForegroundColor Yellow; exit }
        default { if ($cmd -ne "") { Write-Host "[ERROR] Unknown: $cmd (use: go | exit)" -ForegroundColor Red } }
    }
}
