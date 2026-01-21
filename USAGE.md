# Excel.nvim Usage Guide

Complete guide to using excel.nvim for editing Excel files in Neovim.

## Table of Contents

1. [Getting Started](#getting-started)
2. [Basic Operations](#basic-operations)
3. [Advanced Features](#advanced-features)
4. [Working with Formulas](#working-with-formulas)
5. [Tips and Tricks](#tips-and-tricks)
6. [Troubleshooting](#troubleshooting)

## Getting Started

### Opening Your First Excel File

```bash
# From command line
nvim myfile.xlsx

# From within Neovim
:e myfile.xlsx
# or
:ExcelOpen myfile.xlsx
```

### Understanding the Interface

When you open an Excel file, you'll see:

```
   | A                    | B                    | C                    |
---+----------------------+----------------------+----------------------+
  1| Name                 | Age                  | City                 |
  2| John Doe             | 30                   | New York             |
  3| Jane Smith           | 25                   | Los Angeles          |
```

- **Header row**: Shows column letters (A, B, C, etc.)
- **Data rows**: Show row numbers and cell contents
- **Cursor**: Navigate with arrow keys or h/j/k/l

## Basic Operations

### Editing Cells

**Method 1: Direct Edit**
1. Navigate to a cell with arrow keys
2. Press `i`, `a`, or `<CR>` (Enter)
3. Type your content
4. Press `<CR>` to confirm

**Method 2: Command**
```vim
:ExcelEdit
```

### Navigating

```vim
h, j, k, l          " Vim-style navigation
<Arrow keys>        " Standard navigation
gg                  " Go to first cell
G                   " Go to last row
0                   " Go to first column
$                   " Go to last column
ge                  " Go to specific cell (e.g., B5)
```

### Saving Changes

```vim
<leader>w           " Quick save (if configured)
:w                  " Standard Vim save
:ExcelSave          " Explicit Excel save
```

## Advanced Features

### Working with Multiple Sheets

**List All Sheets**
```vim
:ExcelListSheets
# or press: gs
```

**Switch Between Sheets**
```vim
:ExcelSwitchSheet SheetName
# or use: gs (and select from list)
```

**Create New Sheet**
```vim
:ExcelNewSheet MyNewSheet
# or press: gn
```

**Delete Sheet**
```vim
:ExcelDeleteSheet SheetName
# or press: gd (deletes current sheet)
```

### Row and Column Operations

**Insert Row**
```vim
:ExcelInsertRow           " Insert at current position
:ExcelInsertRow 5         " Insert at row 5
# or press: ir
```

**Insert Column**
```vim
:ExcelInsertColumn        " Insert at current position
:ExcelInsertColumn 3      " Insert at column 3
# or press: ic
```

**Delete Row**
```vim
:ExcelDeleteRow           " Delete current row
:ExcelDeleteRow 5         " Delete row 5
# or press: dr
```

**Delete Column**
```vim
:ExcelDeleteColumn        " Delete current column
:ExcelDeleteColumn 3      " Delete column 3
# or press: dc
```

## Working with Formulas

### Basic Formulas

**Insert a Formula**
```vim
:ExcelFormula =SUM(A1:A10)
# or press: gf
```

**Common Formula Examples**
```excel
=SUM(A1:A10)              " Sum a range
=AVERAGE(B1:B20)          " Average
=COUNT(C1:C15)            " Count numbers
=IF(D1>100,"High","Low")  " Conditional
=VLOOKUP(E1,A:B,2,FALSE)  " Lookup
```

### Cross-Sheet Formulas

Reference cells from other sheets:
```excel
=Sheet2!A1               " Single cell
=SUM(Sheet2!A1:A10)      " Range from another sheet
='My Sheet'!B5           " Sheet name with spaces
```

### Recalculating Formulas

After editing formulas or values:
```vim
:ExcelRecalc
# or press: gr
```

**Note:** Requires LibreOffice to be installed.

## Tips and Tricks

### 1. Quick Cell Reference

Jump to any cell quickly:
```vim
:ExcelGoTo B15
:ExcelGoTo AA100
# or press: ge
```

### 2. Batch Operations

Use Vim's powerful editing for repeated operations:
```vim
" Suppose you want to add 10 to column A
" 1. Navigate to first cell in column A
" 2. Start recording macro: qa
" 3. Edit cell, add +10 to formula
" 4. Move down: j
" 5. Stop recording: q
" 6. Repeat 10 times: 10@a
```

### 3. Search in Excel

Use Vim's search functionality:
```vim
/searchterm          " Search forward
?searchterm          " Search backward
n                    " Next match
N                    " Previous match
```

### 4. Copy/Paste Between Cells

```vim
" Copy cell
yy                   " Yank (copy) current row

" Paste cell
p                    " Paste below current cell
P                    " Paste above current cell
```

### 5. Undo/Redo

```vim
u                    " Undo last change
<C-r>                " Redo
```

### 6. Working with CSV Files

CSV files are handled automatically:
```bash
nvim data.csv        " Opens as Excel buffer
```

Edit and save normally - it will remain a CSV file.

## Practical Examples

### Example 1: Creating a Budget

```vim
" 1. Create new Excel file
:ExcelOpen budget.xlsx

" 2. Add headers
" Navigate to A1, press i
Category
" Navigate to B1, press i
Amount

" 3. Add data rows
" A2: Rent, B2: 1200
" A3: Food, B3: 500
" A4: Utilities, B4: 150

" 4. Add total formula
" Navigate to B5, press gf
=SUM(B2:B4)

" 5. Recalculate and save
:ExcelRecalc
:w
```

### Example 2: Grade Calculator

```vim
" Create gradebook with formula
" Columns: Name, Homework, Midterm, Final, Average

" In E2 (Average), press gf:
=(B2*0.3+C2*0.3+D2*0.4)

" Add letter grade in F2, press gf:
=IF(E2>=90,"A",IF(E2>=80,"B",IF(E2>=70,"C","F")))

" Copy formula down
" (Use visual mode and paste, or manual entry)
```

### Example 3: Data Analysis

```vim
" Calculate statistics on data

" Max value: =MAX(A1:A100)
" Min value: =MIN(A1:A100)
" Average: =AVERAGE(A1:A100)
" Count: =COUNT(A1:A100)
" Standard deviation: =STDEV(A1:A100)
```

## Troubleshooting

### Issue: Formulas Show as Text

**Problem:** Formulas appear as `=SUM(A1:A10)` instead of calculated values.

**Solution:**
```vim
:ExcelRecalc
```

**Alternative:** Make sure LibreOffice is installed:
```bash
# Ubuntu/Debian
sudo apt-get install libreoffice

# macOS
brew install libreoffice
```

### Issue: Can't Edit Cells

**Problem:** Pressing `i` doesn't let you edit.

**Possible Causes:**
1. Cursor is on header row or separator
2. File is read-only

**Solution:**
- Navigate to a data cell (row 1+)
- Check file permissions

### Issue: Changes Not Saving

**Problem:** Edits are lost when reopening file.

**Solution:**
- Make sure to save: `:w` or `:ExcelSave`
- Check file permissions
- Verify the file path

### Issue: Python Errors

**Problem:** Errors like "openpyxl not found"

**Solution:**
```bash
pip install openpyxl pandas
```

### Issue: Slow Performance

**Problem:** Large files take long to open/save.

**Solutions:**
1. Use CSV format when possible
2. Split large files into smaller ones
3. Filter data before opening
4. Increase Neovim's memory limits

### Issue: Lost Formatting

**Problem:** Cell colors/formatting not visible.

**Note:** excel.nvim focuses on data editing. Formatting is preserved but not displayed. The formatting will be intact when you open the file in Excel/LibreOffice.

## Keyboard Shortcuts Reference

### Navigation
| Key | Action |
|-----|--------|
| `h/j/k/l` | Move left/down/up/right |
| `ge` | Go to cell (e.g., B5) |
| `gg` | First row |
| `G` | Last row |

### Editing
| Key | Action |
|-----|--------|
| `i / a / <CR>` | Edit cell |
| `gf` | Insert formula |
| `u` | Undo |
| `<C-r>` | Redo |

### Sheets
| Key | Action |
|-----|--------|
| `gs` | List/switch sheets |
| `gn` | New sheet |
| `gd` | Delete sheet |

### Rows/Columns
| Key | Action |
|-----|--------|
| `ir` | Insert row |
| `ic` | Insert column |
| `dr` | Delete row |
| `dc` | Delete column |

### Operations
| Key | Action |
|-----|--------|
| `gr` | Recalculate formulas |
| `<leader>w` | Save file |

## Advanced Configuration

See `examples/advanced_config.lua` for:
- Custom keybindings
- Formula shortcuts
- Auto-save features
- Template creation
- Help window

## Getting Help

```vim
:help excel.nvim           " Plugin help
:ExcelHelp                 " Show keybindings (if configured)
?                          " Quick help (if configured)
```

## Feedback and Contributions

Found a bug? Have a feature request? Please open an issue on GitHub!

---

**Happy Excel editing in Neovim! 🎉**
