-- Advanced configuration example for excel.nvim
-- Save this as ~/.config/nvim/lua/excel_config.lua
-- Then require it in your init.lua: require('excel_config')

local M = {}

-- Custom configuration
M.config = {
  -- Python command
  python_cmd = 'python3',
  
  -- Display settings
  max_col_width = 25,        -- Maximum column width in characters
  min_col_width = 10,        -- Minimum column width
  cell_padding = 2,          -- Padding around cell content
  
  -- Visual settings
  show_gridlines = true,     -- Show cell borders
  show_formulas = false,     -- Show formulas instead of values
  
  -- Behavior
  auto_recalc = false,       -- Auto-recalculate on save (requires LibreOffice)
  
  -- Formatting
  date_format = '%Y-%m-%d',  -- Date display format
  number_format = '%.2f',    -- Number display format (2 decimal places)
}

-- Custom keybindings
M.setup_custom_keybinds = function()
  local excel = require('excel')
  
  -- Quick access commands
  vim.keymap.set('n', '<leader>eo', ':ExcelOpen ', { desc = 'Open Excel file' })
  vim.keymap.set('n', '<leader>es', ':ExcelSave<CR>', { desc = 'Save Excel file' })
  vim.keymap.set('n', '<leader>el', ':ExcelListSheets<CR>', { desc = 'List sheets' })
  vim.keymap.set('n', '<leader>en', ':ExcelNewSheet<CR>', { desc = 'New sheet' })
  vim.keymap.set('n', '<leader>eg', ':ExcelGoTo ', { desc = 'Go to cell' })
  vim.keymap.set('n', '<leader>ef', ':ExcelFormula ', { desc = 'Insert formula' })
  vim.keymap.set('n', '<leader>er', ':ExcelRecalc<CR>', { desc = 'Recalculate formulas' })
  
  -- Quick formula shortcuts
  vim.keymap.set('n', '<leader>sum', function()
    excel.insert_formula('=SUM()')
  end, { desc = 'Insert SUM formula' })
  
  vim.keymap.set('n', '<leader>avg', function()
    excel.insert_formula('=AVERAGE()')
  end, { desc = 'Insert AVERAGE formula' })
  
  vim.keymap.set('n', '<leader>cnt', function()
    excel.insert_formula('=COUNT()')
  end, { desc = 'Insert COUNT formula' })
end

-- Custom commands for common operations
M.setup_custom_commands = function()
  -- Quick financial functions
  vim.api.nvim_create_user_command('ExcelSUM', function(opts)
    local range = opts.args ~= '' and opts.args or 'A1:A10'
    require('excel').insert_formula('=SUM(' .. range .. ')')
  end, { nargs = '?', desc = 'Insert SUM formula with range' })
  
  vim.api.nvim_create_user_command('ExcelAVG', function(opts)
    local range = opts.args ~= '' and opts.args or 'A1:A10'
    require('excel').insert_formula('=AVERAGE(' .. range .. ')')
  end, { nargs = '?', desc = 'Insert AVERAGE formula with range' })
  
  vim.api.nvim_create_user_command('ExcelCOUNT', function(opts)
    local range = opts.args ~= '' and opts.args or 'A1:A10'
    require('excel').insert_formula('=COUNT(' .. range .. ')')
  end, { nargs = '?', desc = 'Insert COUNT formula with range' })
  
  -- Template commands
  vim.api.nvim_create_user_command('ExcelCreateBudget', function()
    vim.notify('Creating budget template...', vim.log.levels.INFO)
    -- You can implement template creation here
  end, { desc = 'Create budget template' })
end

-- Autocommands for Excel files
M.setup_autocommands = function()
  local group = vim.api.nvim_create_augroup('ExcelCustom', { clear = true })
  
  -- Auto-save on focus lost
  vim.api.nvim_create_autocmd('FocusLost', {
    group = group,
    pattern = '*.xlsx',
    callback = function()
      if vim.bo.filetype == 'excel' then
        require('excel').save_excel()
        vim.notify('Excel file auto-saved', vim.log.levels.INFO)
      end
    end,
  })
  
  -- Show notification when opening Excel files
  vim.api.nvim_create_autocmd('FileType', {
    group = group,
    pattern = 'excel',
    callback = function()
      vim.notify('Excel mode active. Press ? for help', vim.log.levels.INFO)
    end,
  })
end

-- Help window for keybindings
M.show_help = function()
  local help_lines = {
    '=== Excel.nvim Help ===',
    '',
    'Navigation:',
    '  <CR> / i / a     Edit cell',
    '  ge               Go to cell',
    '  h/j/k/l          Navigate cells',
    '',
    'Sheets:',
    '  gs               List sheets',
    '  gn               New sheet',
    '  gd               Delete sheet',
    '',
    'Rows/Columns:',
    '  ir               Insert row',
    '  ic               Insert column',
    '  dr               Delete row',
    '  dc               Delete column',
    '',
    'Formulas:',
    '  gf               Insert formula',
    '  gr               Recalculate',
    '  <leader>sum      SUM formula',
    '  <leader>avg      AVERAGE formula',
    '',
    'File:',
    '  <leader>w / :w   Save',
    '  :ExcelSave       Save',
    '',
    'Press q to close this help',
  }
  
  -- Create a floating window
  local buf = vim.api.nvim_create_buf(false, true)
  vim.api.nvim_buf_set_lines(buf, 0, -1, false, help_lines)
  
  local width = 50
  local height = #help_lines
  local row = math.floor((vim.o.lines - height) / 2)
  local col = math.floor((vim.o.columns - width) / 2)
  
  local opts = {
    relative = 'editor',
    width = width,
    height = height,
    row = row,
    col = col,
    style = 'minimal',
    border = 'rounded',
  }
  
  local win = vim.api.nvim_open_win(buf, true, opts)
  
  -- Close on 'q'
  vim.keymap.set('n', 'q', function()
    vim.api.nvim_win_close(win, true)
  end, { buffer = buf })
end

-- Initialize everything
M.setup = function()
  -- Load excel.nvim
  require('excel').setup(M.config)
  
  -- Setup custom features
  M.setup_custom_keybinds()
  M.setup_custom_commands()
  M.setup_autocommands()
  
  -- Add help command
  vim.api.nvim_create_user_command('ExcelHelp', M.show_help, {})
  
  -- Add ? keymap in Excel buffers to show help
  vim.api.nvim_create_autocmd('FileType', {
    pattern = 'excel',
    callback = function()
      vim.keymap.set('n', '?', M.show_help, { buffer = true })
    end,
  })
end

return M
