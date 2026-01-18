-- Excel.nvim - A Neovim plugin for Excel file manipulation
-- Only load once
if vim.g.loaded_excel_nvim then
  return
end
vim.g.loaded_excel_nvim = true

-- Create user commands
vim.api.nvim_create_user_command('ExcelOpen', function(opts)
  require('excel').open(opts.args)
end, { nargs = '?', complete = 'file', desc = 'Open Excel file as CSV in split' })

vim.api.nvim_create_user_command('ExcelView', function(opts)
  require('excel').view(opts.args)
end, { nargs = '?', complete = 'file', desc = 'View Excel file in floating window' })

vim.api.nvim_create_user_command('ExcelEdit', function()
  require('excel').edit()
end, { desc = 'Edit current Excel sheet' })

vim.api.nvim_create_user_command('ExcelSave', function()
  require('excel').save()
end, { desc = 'Save changes to Excel file' })

vim.api.nvim_create_user_command('ExcelCreate', function(opts)
  require('excel').create(opts.args)
end, { nargs = '?', complete = 'file', desc = 'Create new Excel file' })

vim.api.nvim_create_user_command('ExcelSheets', function()
  require('excel').list_sheets()
end, { desc = 'List all sheets in Excel file' })

vim.api.nvim_create_user_command('ExcelSwitchSheet', function(opts)
  require('excel').switch_sheet(opts.args)
end, { nargs = 1, desc = 'Switch to a different sheet' })

vim.api.nvim_create_user_command('ExcelAddFormula', function(opts)
  require('excel.helpers').add_formula(opts.args)
end, { nargs = 1, desc = 'Add formula to current cell' })

vim.api.nvim_create_user_command('ExcelFormat', function()
  require('excel.helpers').format_current()
end, { desc = 'Format current cell/selection' })

vim.api.nvim_create_user_command('ExcelInfo', function()
  require('excel').info()
end, { desc = 'Show Excel file information' })
