# excel.nvim

A Neovim plugin for viewing and editing Excel files (.xlsx) directly within Neovim.

## Features

- 📊 Open and view Excel files as CSV in Neovim
- 📝 Edit Excel data using familiar Neovim interface
- 🔄 Save changes back to Excel format
- 📑 Multiple sheet support (switch between sheets)
- 🪟 Floating window preview
- 🆕 Create new Excel workbooks
- 📋 List all sheets in a workbook
- 🔍 View file information

## Requirements

- Neovim >= 0.7.0
- Python 3 with the following packages:
  - `pandas`
  - `openpyxl`

### Install Python dependencies:

```bash
pip install pandas openpyxl
```

## Installation

### Using [lazy.nvim](https://github.com/folke/lazy.nvim)

```lua
{
  'yourusername/excel.nvim',
  config = function()
    require('excel').setup({
      -- Optional configuration
      python_cmd = 'python3',  -- Python command to use
      auto_recalc = true,       -- Auto-recalculate formulas
      default_sheet = 0,        -- Default sheet index to open
    })
  end,
}
```

### Using [packer.nvim](https://github.com/wbthomason/packer.nvim)

```lua
use {
  'yourusername/excel.nvim',
  config = function()
    require('excel').setup()
  end,
}
```

### Using [vim-plug](https://github.com/junegunn/vim-plug)

```vim
Plug 'yourusername/excel.nvim'

" In your init.vim or init.lua
lua require('excel').setup()
```

## Usage

### Commands

| Command | Description |
|---------|-------------|
| `:ExcelOpen [file]` | Open Excel file as CSV in a split |
| `:ExcelView [file]` | View Excel file in floating window |
| `:ExcelSave` | Save changes back to Excel file |
| `:ExcelCreate [file]` | Create a new Excel workbook |
| `:ExcelSheets` | List all sheets in the current workbook |
| `:ExcelSwitchSheet <index>` | Switch to a different sheet (0-indexed) |
| `:ExcelInfo` | Show information about current Excel file |
| `:ExcelAddFormula <formula>` | Add formula to current cell (planned) |
| `:ExcelFormat` | Format current cell/selection (planned) |

### Example Workflow

1. **Open an Excel file:**
   ```vim
   :ExcelOpen /path/to/your/file.xlsx
   ```

2. **Edit the data** using normal Neovim editing commands

3. **List available sheets:**
   ```vim
   :ExcelSheets
   ```

4. **Switch to another sheet:**
   ```vim
   :ExcelSwitchSheet 1
   ```

5. **Save changes:**
   ```vim
   :ExcelSave
   ```

### Quick Preview

View an Excel file in a floating window without opening it:

```vim
:ExcelView /path/to/file.xlsx
```

Press `q` or `<Esc>` to close the preview.

### Creating a New Workbook

```vim
:ExcelCreate /path/to/new_workbook.xlsx
```

This creates a new Excel file with a default sheet containing three columns.

## Configuration

Default configuration:

```lua
require('excel').setup({
  python_cmd = 'python3',     -- Python command
  temp_dir = vim.fn.stdpath('cache') .. '/excel.nvim',  -- Temp directory
  auto_recalc = true,          -- Auto-recalculate formulas
  default_sheet = 0,           -- Default sheet index
  float_opts = {               -- Floating window options
    relative = 'editor',
    width = math.floor(vim.o.columns * 0.9),
    height = math.floor(vim.o.lines * 0.9),
    col = math.floor(vim.o.columns * 0.05),
    row = math.floor(vim.o.lines * 0.05),
    style = 'minimal',
    border = 'rounded',
  },
})
```

## Key Bindings (Optional)

Add these to your Neovim config for quick access:

```lua
vim.keymap.set('n', '<leader>xo', '<cmd>ExcelOpen<cr>', { desc = 'Open Excel file' })
vim.keymap.set('n', '<leader>xv', '<cmd>ExcelView<cr>', { desc = 'View Excel file' })
vim.keymap.set('n', '<leader>xs', '<cmd>ExcelSave<cr>', { desc = 'Save Excel file' })
vim.keymap.set('n', '<leader>xl', '<cmd>ExcelSheets<cr>', { desc = 'List sheets' })
vim.keymap.set('n', '<leader>xi', '<cmd>ExcelInfo<cr>', { desc = 'Excel info' })
```

## Recommended Plugins

For better CSV editing experience, consider installing:

- [chrisbra/csv.vim](https://github.com/chrisbra/csv.vim) - CSV file handling and column alignment
- [mechatroner/rainbow_csv](https://github.com/mechatroner/rainbow_csv) - CSV syntax highlighting

## How It Works

1. **Excel → CSV**: When you open an Excel file, the plugin uses `pandas` to convert the active sheet to a temporary CSV file
2. **Edit**: You edit the CSV file using Neovim's standard editing features
3. **CSV → Excel**: When you save, `openpyxl` updates the original Excel file with your changes

## Limitations

- Formulas are preserved when saving, but you edit the calculated values in CSV format
- Complex Excel formatting may not be fully preserved
- Charts and images are not supported
- Macros are not executed

## Future Enhancements

- [ ] Direct formula editing
- [ ] Cell formatting (bold, colors, etc.)
- [ ] Better formula support
- [ ] Column width adjustment
- [ ] Row/column insertion/deletion
- [ ] Cell merging
- [ ] Data validation
- [ ] Better error handling
- [ ] Undo/redo support for Excel operations

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - see LICENSE file for details

## Credits

Built with:
- [pandas](https://pandas.pydata.org/) - Data manipulation
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel file handling
- [Neovim](https://neovim.io/) - The best text editor

## Support

If you encounter any issues or have feature requests, please open an issue on GitHub.
