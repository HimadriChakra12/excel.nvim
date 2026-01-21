-- excel.nvim - Full-featured Excel editor for Neovim
-- Supports .xls, .xlsx, .xlsm, .xlsb formats

if vim.g.loaded_excel_nvim then
  return
end
vim.g.loaded_excel_nvim = 1

-- Set up autocommands for Excel files
vim.api.nvim_create_augroup('ExcelNvim', { clear = true })

-- Auto-detect and open Excel files
vim.api.nvim_create_autocmd({ 'BufReadCmd' }, {
  group = 'ExcelNvim',
  pattern = { '*.xlsx', '*.xls', '*.xlsm', '*.xlsb', '*.csv' },
  callback = function(args)
    require('excel').open_excel(args.file)
  end,
})

-- Save Excel files
vim.api.nvim_create_autocmd({ 'BufWriteCmd' }, {
  group = 'ExcelNvim',
  pattern = { '*.xlsx', '*.xls', '*.xlsm', '*.xlsb', '*.csv' },
  callback = function(args)
    require('excel').save_excel(args.file)
  end,
})

-- Commands
vim.api.nvim_create_user_command('ExcelOpen', function(opts)
  require('excel').open_excel(opts.args)
end, { nargs = 1, complete = 'file' })

vim.api.nvim_create_user_command('ExcelSave', function()
  require('excel').save_excel()
end, {})

vim.api.nvim_create_user_command('ExcelNewSheet', function(opts)
  require('excel').new_sheet(opts.args)
end, { nargs = '?' })

vim.api.nvim_create_user_command('ExcelDeleteSheet', function(opts)
  require('excel').delete_sheet(opts.args)
end, { nargs = '?' })

vim.api.nvim_create_user_command('ExcelListSheets', function()
  require('excel').list_sheets()
end, {})

vim.api.nvim_create_user_command('ExcelSwitchSheet', function(opts)
  require('excel').switch_sheet(opts.args)
end, { nargs = 1 })

vim.api.nvim_create_user_command('ExcelInsertRow', function(opts)
  require('excel').insert_row(opts.args)
end, { nargs = '?' })

vim.api.nvim_create_user_command('ExcelInsertColumn', function(opts)
  require('excel').insert_column(opts.args)
end, { nargs = '?' })

vim.api.nvim_create_user_command('ExcelDeleteRow', function(opts)
  require('excel').delete_row(opts.args)
end, { nargs = '?' })

vim.api.nvim_create_user_command('ExcelDeleteColumn', function(opts)
  require('excel').delete_column(opts.args)
end, { nargs = '?' })

vim.api.nvim_create_user_command('ExcelFormat', function()
  require('excel').format_cell()
end, {})

vim.api.nvim_create_user_command('ExcelFormula', function(opts)
  require('excel').insert_formula(opts.args)
end, { nargs = 1 })

vim.api.nvim_create_user_command('ExcelRecalc', function()
  require('excel').recalculate()
end, {})

vim.api.nvim_create_user_command('ExcelGoTo', function(opts)
  require('excel').goto_cell(opts.args)
end, { nargs = 1 })

vim.api.nvim_create_user_command('ExcelFreeze', function(opts)
  require('excel').freeze_panes(opts.args)
end, { nargs = '?' })

vim.api.nvim_create_user_command('ExcelSort', function(opts)
  require('excel').sort_range(opts.args)
end, { nargs = '?' })

vim.api.nvim_create_user_command('ExcelFilter', function()
  require('excel').toggle_filter()
end, {})

vim.api.nvim_create_user_command('ExcelChart', function(opts)
  require('excel').create_chart(opts.args)
end, { nargs = '*' })
