# Quick Start Guide - excel.nvim

Get up and running with excel.nvim in 5 minutes!

## Installation (1 minute)

### Step 1: Install Python Dependencies
```bash
pip install openpyxl pandas
```

### Step 2: Run Install Script
```bash
cd excel.nvim
./install.sh
```

### Step 3: Add to Neovim Config

Add to your `~/.config/nvim/init.lua`:

```lua
-- Using lazy.nvim
{
  'excel.nvim',
  dir = '/path/to/excel.nvim',
  config = function()
    require('excel').setup()
  end,
}

-- Or manually
vim.opt.runtimepath:append('/path/to/excel.nvim')
```

Restart Neovim.

## First Use (2 minutes)

### 1. Create Sample Files
```bash
cd excel.nvim
python3 create_samples.py
```

### 2. Open an Excel File
```bash
nvim sample_budget.xlsx
```

### 3. Try These Actions

**Edit a cell:**
- Navigate with arrow keys to any cell
- Press `i` or Enter
- Type a new value
- Press Enter to confirm

**Add a formula:**
- Navigate to an empty cell
- Press `gf`
- Type: `=SUM(B2:B6)`
- Press Enter

**Switch sheets:**
- Press `gs`
- Select a sheet from the list

**Save your work:**
- Press `:w` and Enter

## Essential Keybindings (2 minutes)

Remember these 10 keys and you're 80% there:

```
i          Edit cell
gs         Switch sheets
ge         Go to cell (like B5)
gf         Add formula
ir/ic      Insert row/column
dr/dc      Delete row/column
:w         Save
u          Undo
gr         Recalculate formulas
?          Help (if configured)
```

## Common Tasks

### Create a New Excel File
```bash
nvim new_file.xlsx
```
Start editing! File will be created on save.

### Convert CSV to Excel
```bash
nvim data.csv
:ExcelSave data.xlsx
```

### Add Calculations
1. Navigate to cell → Press `gf`
2. Type formula: `=A1+B1` or `=SUM(A1:A10)`
3. Press Enter
4. Press `gr` to recalculate

## What's Next?

- Read `USAGE.md` for detailed guide
- Check `examples/advanced_config.lua` for customization
- See `README.md` for all features

## Need Help?

Common issues and solutions:

**"openpyxl not found"**
```bash
pip install openpyxl
```

**Formulas show as text**
```bash
# Install LibreOffice for formula calculation
sudo apt-get install libreoffice  # Ubuntu/Debian
brew install libreoffice          # macOS
```

**Changes not saving**
- Make sure to press `:w` or `:ExcelSave`

## Tips for Vim Users

- All standard Vim navigation works (gg, G, 0, $)
- Use visual mode for selections
- Macros work great for repetitive tasks
- Use `/` to search in your spreadsheet
- `:set number` shows Vim line numbers (optional)

## Tips for Excel Users

| Excel | Neovim |
|-------|--------|
| Click cell → Type | Navigate → `i` → Type |
| Ctrl+S | `:w` |
| F2 | `i` |
| Ctrl+Z | `u` |
| Sheet tabs | `gs` |
| Ctrl+Home | `gg` |
| Ctrl+End | `G` |

---

**You're ready to edit Excel files in Neovim! 🚀**

For more details, see:
- `USAGE.md` - Complete usage guide
- `README.md` - Full documentation
- `:help excel` - In-editor help
