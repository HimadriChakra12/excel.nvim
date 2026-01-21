# excel.nvim

A full-featured Excel editor for Neovim that allows you to view, edit, and modify Excel files directly in your editor.

> I Need more help to get at this perfectly. Can Help me if want to support this thing.

## Features

✅ **Multiple Format Support**
- `.xlsx` - Excel 2007+ (full support)
- `.xls` - Excel 97-2003 (read and convert to .xlsx)
- `.xlsm` - Excel Macro-Enabled
- `.xlsb` - Excel Binary
- `.csv` - Comma Separated Values

✅ **Core Functionality**
- View Excel files in a grid layout
- Edit cell values with formula support
- Multiple sheet management
- Insert/delete rows and columns
- Navigate between cells easily
- Save changes back to Excel format
- Formula recalculation (with LibreOffice)

✅ **Advanced Features**
- Syntax highlighting for formulas and numbers
- Cell navigation with intuitive keybindings
- Sheet switching and management
- Go-to cell functionality
- CSV import/export

## Requirements

### Required
- Neovim >= 0.8.0
- Python 3.7+
- `openpyxl` Python package

### Optional
- `pandas` (for better CSV handling and data analysis)
- LibreOffice (for formula recalculation)

## Installation

### 1. Install Python Dependencies

```bash
pip install openpyxl pandas
```

### 2. Install Plugin

#### Using [lazy.nvim](https://github.com/folke/lazy.nvim)

```lua
{
  'excel.nvim',
  dir = '/path/to/excel.nvim',
  config = function()
    require('excel').setup({
      -- Configuration options
      python_cmd = 'python3',
      max_col_width = 20,
      min_col_width = 8,
      show_gridlines = true,
      auto_recalc = true,
    })
  end,
}
```

#### Using [packer.nvim](https://github.com/wbthomason/packer.nvim)

```lua
use {
  'excel.nvim',
  config = function()
    require('excel').setup()
  end
}
```

#### Using [vim-plug](https://github.com/junegunn/vim-plug)

```vim
Plug '/path/to/excel.nvim'
```

#### Manual Installation

```bash
# Clone or copy the plugin to your Neovim config directory
mkdir -p ~/.local/share/nvim/site/pack/plugins/start/
cp -r excel.nvim ~/.local/share/nvim/site/pack/plugins/start/
```

### 3. Install LibreOffice (Optional, for formula recalculation)

**Ubuntu/Debian:**
```bash
sudo apt-get install libreoffice
```

**macOS:**
```bash
brew install libreoffice
```

**Windows:**
Download from [LibreOffice website](https://www.libreoffice.org/download/download/)

## Usage

### Opening Excel Files

Excel files are automatically opened when you open them in Neovim:

```bash
nvim myfile.xlsx
```

Or from within Neovim:

```vim
:ExcelOpen myfile.xlsx
```

### Keybindings

#### Navigation
- `<CR>` / `i` / `a` - Edit cell at cursor
- `ge` - Go to specific cell (e.g., B5)
- `h/j/k/l` - Navigate cells (Vim navigation)

#### Sheet Management
- `gs` - List and switch between sheets
- `gn` - Create new sheet
- `gd` - Delete current sheet

#### Row/Column Operations
- `ir` - Insert row at cursor
- `ic` - Insert column at cursor
- `dr` - Delete row at cursor
- `dc` - Delete column at cursor

#### Formula Operations
- `gf` - Insert formula in current cell
- `gr` - Recalculate all formulas (requires LibreOffice)

#### File Operations
- `<leader>w` or `:w` - Save Excel file
- `:ExcelSave` - Save Excel file

### Commands

#### File Operations
```vim
:ExcelOpen <filepath>        " Open Excel file
:ExcelSave                    " Save current file
```

#### Sheet Operations
```vim
:ExcelListSheets             " Show list of sheets
:ExcelSwitchSheet <name>     " Switch to specific sheet
:ExcelNewSheet [name]        " Create new sheet
:ExcelDeleteSheet [name]     " Delete sheet
```

#### Cell/Row/Column Operations
```vim
:ExcelInsertRow [position]   " Insert row at position
:ExcelInsertColumn [position] " Insert column at position
:ExcelDeleteRow [position]   " Delete row at position
:ExcelDeleteColumn [position] " Delete column at position
```

#### Formula Operations
```vim
:ExcelFormula <formula>      " Insert formula (e.g., =SUM(A1:A10))
:ExcelRecalc                 " Recalculate all formulas
```

#### Navigation
```vim
:ExcelGoTo <cell>            " Go to cell (e.g., B5, AA100)
```

#### Future Features (Coming Soon)
```vim
:ExcelFormat                 " Format cell (bold, colors, etc.)
:ExcelFreeze [position]      " Freeze panes
:ExcelSort [options]         " Sort range
:ExcelFilter                 " Toggle autofilter
:ExcelChart [options]        " Create chart
```

## Examples

### Example 1: Edit a Cell

1. Open an Excel file: `nvim budget.xlsx`
2. Navigate to a cell with arrow keys
3. Press `i` or `<CR>` to edit
4. Type the new value
5. Press `<CR>` to confirm
6. Save with `<leader>w`

### Example 2: Add a Formula

1. Navigate to the cell where you want the formula
2. Press `gf`
3. Enter the formula: `=SUM(A1:A10)`
4. Press `<CR>` to confirm
5. Press `gr` to recalculate (requires LibreOffice)

### Example 3: Work with Multiple Sheets

1. Press `gs` to see all sheets
2. Select a sheet from the list
3. Press `gn` to create a new sheet
4. Enter the sheet name
5. Switch back with `gs`

### Example 4: Manipulate Rows/Columns

1. Navigate to a row/column
2. Press `ir` to insert a row above
3. Press `ic` to insert a column to the left
4. Press `dr` to delete the current row
5. Press `dc` to delete the current column

## Configuration

Configure the plugin in your `init.lua`:

```lua
require('excel').setup({
  -- Python command (python3, python, etc.)
  python_cmd = 'python3',
  
  -- Column display width
  max_col_width = 20,
  min_col_width = 8,
  
  -- Display options
  show_gridlines = true,
  show_formulas = false,  -- Show formulas instead of values
  
  -- Auto-recalculate on save (requires LibreOffice)
  auto_recalc = false,
  
  -- Formatting
  date_format = '%Y-%m-%d',
  number_format = '%.2f',
  cell_padding = 1,
})
```

## How It Works

1. **Python Backend**: Uses `openpyxl` to read/write Excel files
2. **Lua Frontend**: Provides Neovim interface and keybindings
3. **Data Format**: Converts Excel to a grid format for editing
4. **Formula Support**: Preserves formulas and recalculates with LibreOffice
5. **Multi-format**: Handles different Excel formats through conversion

## Architecture

```
excel.nvim/
├── plugin/
│   └── excel.lua          # Plugin initialization & autocommands
├── lua/
│   └── excel.lua          # Core Lua module
├── python/
│   └── excel_handler.py   # Python Excel handler
├── syntax/
│   └── excel.vim          # Syntax highlighting
├── ftdetect/
│   └── excel.vim          # Filetype detection
└── README.md              # This file
```

## Troubleshooting

### "openpyxl not installed"
```bash
pip install openpyxl
```

### "Formula recalculation failed"
Install LibreOffice:
```bash
# Ubuntu/Debian
sudo apt-get install libreoffice

# macOS
brew install libreoffice
```

### "Python command not found"
Configure the Python command in setup:
```lua
require('excel').setup({
  python_cmd = '/usr/bin/python3',  -- or your Python path
})
```

### Formulas show as text
Press `gr` to recalculate formulas (requires LibreOffice)

### Can't edit cells
Make sure you're not on the header row or separator line. Navigate to a data cell and press `i`.

## Performance

- **Small files** (< 1000 rows): Instant loading
- **Medium files** (< 10000 rows): A few seconds
- **Large files** (> 10000 rows): May take 10-30 seconds

For very large files, consider:
- Using CSV format when possible
- Filtering/limiting the data before opening
- Using external tools for initial processing

## Limitations

- Formula recalculation requires LibreOffice
- Charts and images are not displayed (preserved on save)
- Advanced formatting (colors, fonts) shown simply
- Macros in .xlsm files are preserved but not executed
- Very large files (>50MB) may be slow

## Roadmap

- [ ] Cell formatting (bold, colors, borders)
- [ ] Freeze panes visualization
- [ ] Sort and filter UI
- [ ] Chart creation and editing
- [ ] Better performance for large files
- [ ] Conditional formatting support
- [ ] Data validation
- [ ] Pivot table support
- [ ] VBA macro editing

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

## License

MIT License - See LICENSE file for details

## Credits

- Built with [openpyxl](https://openpyxl.readthedocs.io/)
- Inspired by various Vim/Neovim table plugins
- Created for the Neovim community

## Support

If you encounter any issues or have feature requests, please open an issue on GitHub.

---

**Made with ❤️ for Neovim users who work with Excel files**
