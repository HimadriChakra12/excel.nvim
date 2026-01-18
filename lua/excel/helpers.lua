-- Excel.nvim helpers for formulas and formatting
local M = {}

-- Add formula to current cell (requires coordinates)
function M.add_formula(formula)
  local excel = require('excel')
  
  if not excel.state.original_file then
    vim.notify('No Excel file is currently open', vim.log.levels.WARN)
    return
  end
  
  -- Get current cursor position
  local cursor = vim.api.nvim_win_get_cursor(0)
  local row = cursor[1]
  local col = cursor[2]
  
  -- Convert to Excel coordinates (1-indexed)
  local excel_row = row
  local excel_col = col + 1
  
  -- Convert column number to Excel letter (A, B, C, ...)
  local function col_to_letter(n)
    local result = ''
    while n > 0 do
      local remainder = (n - 1) % 26
      result = string.char(65 + remainder) .. result
      n = math.floor((n - 1) / 26)
    end
    return result
  end
  
  local cell_ref = col_to_letter(excel_col) .. excel_row
  
  vim.notify('Formula will be added to cell ' .. cell_ref .. ' when saved', vim.log.levels.INFO)
  vim.notify('Note: Direct formula editing not yet implemented. Use ExcelSave to update.', vim.log.levels.WARN)
end

-- Format current cell (basic implementation)
function M.format_current()
  vim.notify('Formatting features coming soon!', vim.log.levels.INFO)
  vim.notify('For now, edit formulas and formatting directly in the Excel file', vim.log.levels.INFO)
end

-- CSV utilities
M.csv = {}

function M.csv.align_columns()
  local buf = vim.api.nvim_get_current_buf()
  if vim.bo[buf].filetype ~= 'csv' then
    vim.notify('Not a CSV buffer', vim.log.levels.WARN)
    return
  end
  
  -- Simple column alignment for CSV viewing
  vim.notify('Use a CSV plugin like chrisbra/csv.vim for advanced formatting', vim.log.levels.INFO)
end

return M
