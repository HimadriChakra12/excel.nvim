-- Example configuration for excel.nvim
-- Add this to your Neovim config (init.lua)

return {
  'yourusername/excel.nvim',
  dependencies = {
    -- Optional: Better CSV editing experience
    'chrisbra/csv.vim',
    -- Optional: CSV syntax highlighting
    'mechatroner/rainbow_csv',
  },
  config = function()
    require('excel').setup({
      -- Python command to use (default: 'python3')
      python_cmd = 'python3',
      
      -- Temporary directory for CSV files
      temp_dir = vim.fn.stdpath('cache') .. '/excel.nvim',
      
      -- Auto-recalculate formulas when saving
      auto_recalc = true,
      
      -- Default sheet index to open (0-indexed)
      default_sheet = 0,
      
      -- Floating window configuration
      float_opts = {
        relative = 'editor',
        width = math.floor(vim.o.columns * 0.9),
        height = math.floor(vim.o.lines * 0.9),
        col = math.floor(vim.o.columns * 0.05),
        row = math.floor(vim.o.lines * 0.05),
        style = 'minimal',
        border = 'rounded',
      },
    })
    
    -- Optional: Set up key mappings
    local map = vim.keymap.set
    
    -- Excel commands
    map('n', '<leader>xo', '<cmd>ExcelOpen<cr>', { desc = 'Excel: Open file' })
    map('n', '<leader>xv', '<cmd>ExcelView<cr>', { desc = 'Excel: View file' })
    map('n', '<leader>xs', '<cmd>ExcelSave<cr>', { desc = 'Excel: Save file' })
    map('n', '<leader>xc', '<cmd>ExcelCreate<cr>', { desc = 'Excel: Create new file' })
    map('n', '<leader>xl', '<cmd>ExcelSheets<cr>', { desc = 'Excel: List sheets' })
    map('n', '<leader>xi', '<cmd>ExcelInfo<cr>', { desc = 'Excel: Show info' })
    
    -- Sheet navigation
    map('n', '<leader>x1', '<cmd>ExcelSwitchSheet 0<cr>', { desc = 'Excel: Switch to sheet 0' })
    map('n', '<leader>x2', '<cmd>ExcelSwitchSheet 1<cr>', { desc = 'Excel: Switch to sheet 1' })
    map('n', '<leader>x3', '<cmd>ExcelSwitchSheet 2<cr>', { desc = 'Excel: Switch to sheet 2' })
    
    -- Auto-open Excel files
    vim.api.nvim_create_autocmd('BufReadCmd', {
      pattern = '*.xlsx',
      callback = function()
        require('excel').open(vim.fn.expand('<afile>'))
      end,
    })
  end,
}
