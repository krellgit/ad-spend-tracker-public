# Amazon Advertising Spend Tracker

Advanced work simulator with 145+ micro-tasks and real file operations for Amazon PPC campaign management workflows.

## Features

### v4.0 - BULK Edition

1. **145 Micro-Tasks**
   - 55 Excel operations (data entry, formulas, navigation)
   - 40 Perpetua UI interactions (goals, streams, analytics)
   - 20 Chrome/web operations
   - 10 Teams communications
   - 5 Human behaviors (pauses, fidgets, re-reads)
   - **15 NEW bulk file operations** (filter, sort, navigate 100k+ rows)

2. **Real File Operations**
   - Opens and operates on actual 1.2GB Excel file
   - Authentic filtering, sorting, navigation on massive datasets
   - Real Excel COM automation (not just keystrokes)
   - Genuine performance characteristics (lag, memory usage)

3. **Anti-Pattern Detection**
   - 20-task memory prevents repetition
   - Time-of-day category weighting
   - Weighted randomization (6.78 bits entropy)
   - Human behavior injection
   - Variable task durations and gaps

4. **Perpetua UI Precision**
   - SP/SB/SD Goals navigation
   - Streams management
   - Analytics and reporting
   - Realistic UI coordinates

## Installation

### Global Install (Recommended)

```bash
npm install -g @krellgit/ad-spend-tracker
```

Then run from anywhere:

```bash
ast                # Standard version (130 tasks)
ast bulk           # BULK edition (145 tasks + real file ops)
```

### Local Install

```bash
npm install @krellgit/ad-spend-tracker
```

Run via npx:

```bash
npx ad-spend-tracker
npx ad-spend-tracker bulk
```

## Usage

### Standard Version

```bash
ad-spend-tracker
# or
ast
```

**Features:**
- 130 micro-tasks
- Simulated data entry
- Time-of-day awareness
- 20-task anti-pattern memory

### BULK Edition

```bash
ad-spend-tracker bulk
# or
ast bulk
```

**Additional Features:**
- 145 tasks (15 new bulk operations)
- Opens "Bulk sample.xlsx" (1.2GB)
- Real Excel COM automation
- Operates on 100k+ rows
- Authentic file I/O footprint

**Requirements for BULK mode:**
- Place `Bulk sample.xlsx` in project directory
- File size: ~1.2GB
- Contains real campaign data

### Command-Line Options

```bash
ast --help          # Show help
ast --version       # Show version
ast bulk            # Run BULK edition
```

## Setup

1. **Install globally:**
   ```bash
   npm install -g @krellgit/ad-spend-tracker
   ```

2. **For BULK mode, prepare data file:**
   - Ensure `Bulk sample.xlsx` is accessible
   - Place in project root or working directory
   - Verify file size (~1.2GB)

3. **Open required applications:**
   - Excel (for Excel tasks)
   - Chrome with Perpetua app (for Perpetua tasks)
   - Teams (for Teams tasks)

4. **Run the tracker:**
   ```bash
   ast              # Standard
   ast bulk         # BULK edition
   ```

5. **In the PowerShell interface:**
   ```
   tracker> go      # Start simulation
   tracker> exit    # Stop and close
   ```

## How It Works

### Task Selection (Anti-Pattern)

1. **Weighted Pool**
   - Each task has category weight and individual weight
   - Total weight = task_weight × (category_weight / 10)
   - Higher weights = more frequent selection

2. **20-Task Memory**
   - Maintains history of last 20 executed tasks
   - Excludes these from selection pool
   - Prevents short-term repetition patterns
   - Ensures 110+ tasks always available

3. **Time-of-Day Adaptation**
   - **Morning (9-12):** Heavy Perpetua (35%) + Excel (30%)
   - **Lunch (12-14):** Balanced, lighter activity
   - **Afternoon (14-17):** Heavy Excel (40%) + optimization

### BULK File Operations

When running in BULK mode:

1. **Startup**
   - Opens Excel COM object
   - Loads "Bulk sample.xlsx" (1.2GB)
   - Analyzes structure (row count, worksheets)
   - Creates navigation ranges

2. **Real Operations**
   - Jump to specific rows (50k, 100k)
   - Filter columns with real criteria
   - Sort by ACOS, Spend, etc.
   - Search for actual SKUs/campaigns
   - Navigate between worksheets
   - Copy large data ranges

3. **Performance Characteristics**
   - Real Excel rendering lag on 100k+ rows
   - Authentic filter/sort processing time
   - 1-2GB memory footprint
   - Genuine file I/O operations

## Task Categories

### Excel (55 tasks)
- Data entry: SKUs, campaigns, ACOS, spend, bids
- Formulas: SUM, AVERAGE, ROAS calculations
- Navigation: scroll, click, select, tabs
- Operations: copy, paste, save, find, undo

### Bulk (15 tasks - BULK mode only)
- Jump to row 50,000 / 100,000
- Filter ACOS > 20%
- Sort by Spend / ACOS
- Search specific SKU in bulk data
- Navigate worksheets
- Review large ranges
- Copy bulk data

### Perpetua (40 tasks)
- Goals: SP/SB/SD navigation, metrics review
- Streams: view, filter, expand, automation
- General: sidebar, search, date range, refresh

### Chrome (20 tasks)
- Campaign manager, ad groups, keywords
- Search terms, reports, exports
- Google searches, article reading

### Teams (10 tasks)
- Chat threads, messages, channels
- Activity feed, files, mentions

### Human (5 tasks)
- Pauses, fidgets, hesitations
- Re-reading, micro-breaks

## Configuration

### Anti-Pattern Settings

Located in script header:

```powershell
$script:taskHistory = @()  # Last 20 tasks
$antiPatternMemory = 20    # Exclusion count
```

**Current optimal: 20 tasks**
- 145 total tasks → 125 available at any moment
- 13.8% exclusion rate
- 6.86 bits entropy per selection

### Category Weights

Modify `Get-CategoryWeights` function:

```powershell
function Get-CategoryWeights {
    $hour = (Get-Date).Hour
    if ($hour -ge 9 -and $hour -lt 12) {
        return @{
            excel=25;
            bulk=15;      # Morning: review bulk data
            perpetua=30;
            chrome=15;
            teams=10;
            human=5
        }
    }
    # ... more time periods
}
```

## NPM Deployment

### Publish to NPM

1. **Login to NPM:**
   ```bash
   npm login
   ```

2. **Publish package:**
   ```bash
   cd /path/to/rmm-research
   npm publish --access public
   ```

3. **Update version:**
   ```bash
   npm version patch   # 4.0.0 → 4.0.1
   npm version minor   # 4.0.1 → 4.1.0
   npm version major   # 4.1.0 → 5.0.0
   npm publish
   ```

### GitHub Repository

```bash
git remote add origin https://github.com/krellgit/rmm-research.git
git push -u origin main
```

## Requirements

- **OS:** Windows 10/11 (win32)
- **Node.js:** >= 14.0.0
- **PowerShell:** 5.1+
- **Excel:** Microsoft Excel 2016+ (for COM automation)
- **Bulk file:** Bulk sample.xlsx (1.2GB) for BULK mode

## Development

### Local Development

```bash
git clone https://github.com/krellgit/rmm-research.git
cd rmm-research
npm install
node bin/cli.js         # Test standard version
node bin/cli.js bulk    # Test BULK edition
```

### File Structure

```
rmm-research/
├── bin/
│   └── cli.js                        # NPM CLI wrapper
├── scripts/
│   ├── ad_spend_tracker.ps1          # v3.1 (130 tasks)
│   └── ad_spend_tracker_bulk.ps1     # v4.0 (145 tasks + bulk ops)
├── Bulk sample.xlsx                  # 1.2GB bulk data file
├── package.json
└── README.md
```

## Troubleshooting

### "Bulk sample.xlsx not found"

Place the 1.2GB file in the project root:
```
rmm-research/Bulk sample.xlsx
```

### "This tool only runs on Windows"

Package is Windows-only. Requires win32 platform and PowerShell.

### PowerShell Execution Policy

If script fails to run:
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser
```

### Excel COM Errors

Ensure Excel is installed and not running in protected mode.

## License

MIT License - see LICENSE file

## Author

**krellgit**
- GitHub: [@krellgit](https://github.com/krellgit)
- Repository: [rmm-research](https://github.com/krellgit/rmm-research)

## Changelog

### v4.0.0 - BULK Edition
- Added 15 bulk file operation tasks (145 total)
- Real Excel COM automation for 1.2GB file
- Authentic filter/sort/navigate on 100k+ rows
- Enhanced NPM deployment support
- CLI wrapper with --help and --version

### v3.1.0
- 130 micro-tasks
- 20-task anti-pattern memory
- Time-of-day category weighting
- Perpetua UI precision
- Human behavior injection

### v3.0.0
- Initial NPM package
- Basic work simulation
- Excel and Chrome tasks
