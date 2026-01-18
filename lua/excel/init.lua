-- Excel.nvim main module
local M = {}
local uv = vim.loop

-- Plugin state
M.state = {
  current_file = nil,
  current_sheet = nil,
  original_file = nil,
  temp_csv = nil,
  sheets = {},
  excel_buf = nil,
}

-- Configuration
M.config = {
  python_cmd = 'python3',
  temp_dir = vim.fn.stdpath('cache') .. '/excel.nvim',
  auto_recalc = true,
  default_sheet = 0,
  float_opts = {
    relative = 'editor',
    width = math.floor(vim.o.columns * 0.9),
    height = math.floor(vim.o.lines * 0.9),
    col = math.floor(vim.o.columns * 0.05),
    row = math.floor(vim.o.lines * 0.05),
    style = 'minimal',
    border = 'rounded',
  },
}

-- Setup function for user configuration
function M.setup(opts)
  M.config = vim.tbl_deep_extend('force', M.config, opts or {})
  -- Ensure temp directory exists
  vim.fn.mkdir(M.config.temp_dir, 'p')
end

-- Utility: Check if file exists
local function file_exists(path)
  local stat = uv.fs_stat(path)
  return stat ~= nil and stat.type == 'file'
end

-- Utility: Get absolute path
local function get_absolute_path(path)
  if not path or path == '' then
    return nil
  end
  return vim.fn.fnamemodify(path, ':p')
end

-- Utility: Run Python script
local function run_python(script, callback)
  local stdout = {}
  local stderr = {}
  
  local handle
  handle = uv.spawn(M.config.python_cmd, {
    args = { '-c', script },
    stdio = { nil, uv.new_pipe(false), uv.new_pipe(false) },
  }, function(code, signal)
    if callback then
      callback(code, table.concat(stdout), table.concat(stderr))
    end
  end)
  
  if not handle then
    vim.notify('Failed to spawn Python process', vim.log.levels.ERROR)
    return
  end
  
  uv.read_start(handle.stdio[2], function(err, data)
    if data then table.insert(stdout, data) end
  end)
  
  uv.read_start(handle.stdio[3], function(err, data)
    if data then table.insert(stderr, data) end
  end)
end

-- Convert Excel to CSV
function M.to_csv(excel_file, sheet_index, callback)
  sheet_index = sheet_index or 0
  local csv_file = M.config.temp_dir .. '/' .. vim.fn.fnamemodify(excel_file, ':t:r') .. '_sheet' .. sheet_index .. '.csv'
  
  local script = string.format([[
import pandas as pd
import sys

try:
    xlsx = pd.read_excel('%s', sheet_name=%d)
    xlsx.to_csv('%s', index=False)
    print('SUCCESS')
except Exception as e:
    print(f'ERROR: {str(e)}', file=sys.stderr)
    sys.exit(1)
]], excel_file, sheet_index, csv_file)
  
  run_python(script, function(code, stdout, stderr)
    if code == 0 then
      callback(csv_file, nil)
    else
      callback(nil, stderr)
    end
  end)
end

-- Get sheet names
function M.get_sheets(excel_file, callback)
  local script = string.format([[
import pandas as pd
import json
import sys

try:
    xlsx = pd.ExcelFile('%s')
    sheets = xlsx.sheet_names
    print(json.dumps(sheets))
except Exception as e:
    print(f'ERROR: {str(e)}', file=sys.stderr)
    sys.exit(1)
]], excel_file)
  
  run_python(script, function(code, stdout, stderr)
    if code == 0 then
      local sheets = vim.fn.json_decode(stdout)
      callback(sheets, nil)
    else
      callback(nil, stderr)
    end
  end)
end

-- Open Excel file
function M.open(file)
  local excel_file = get_absolute_path(file or vim.fn.expand('%'))
  
  if not excel_file or not file_exists(excel_file) then
    vim.notify('File does not exist: ' .. (excel_file or 'nil'), vim.log.levels.ERROR)
    return
  end
  
  M.state.original_file = excel_file
  
  M.get_sheets(excel_file, function(sheets, err)
    if err then
      vim.notify('Error getting sheets: ' .. err, vim.log.levels.ERROR)
      return
    end
    
    M.state.sheets = sheets
    M.state.current_sheet = M.config.default_sheet
    
    M.to_csv(excel_file, M.state.current_sheet, function(csv_file, err)
      if err then
        vim.notify('Error converting to CSV: ' .. err, vim.log.levels.ERROR)
        return
      end
      
      M.state.temp_csv = csv_file
      vim.cmd('split ' .. vim.fn.fnameescape(csv_file))
      vim.bo.filetype = 'csv'
      M.state.excel_buf = vim.api.nvim_get_current_buf()
      
      -- Set buffer name to show Excel file and sheet
      local sheet_name = sheets[M.state.current_sheet + 1] or 'Sheet' .. M.state.current_sheet
      vim.api.nvim_buf_set_name(M.state.excel_buf, 
        excel_file .. ' [' .. sheet_name .. ']')
      
      vim.notify('Opened: ' .. sheet_name, vim.log.levels.INFO)
    end)
  end)
end

-- View in floating window
function M.view(file)
  local excel_file = get_absolute_path(file or vim.fn.expand('%'))
  
  if not excel_file or not file_exists(excel_file) then
    vim.notify('File does not exist: ' .. (excel_file or 'nil'), vim.log.levels.ERROR)
    return
  end
  
  M.to_csv(excel_file, M.config.default_sheet, function(csv_file, err)
    if err then
      vim.notify('Error converting to CSV: ' .. err, vim.log.levels.ERROR)
      return
    end
    
    local buf = vim.api.nvim_create_buf(false, true)
    local lines = vim.fn.readfile(csv_file)
    vim.api.nvim_buf_set_lines(buf, 0, -1, false, lines)
    vim.bo[buf].filetype = 'csv'
    vim.bo[buf].modifiable = false
    
    local win = vim.api.nvim_open_win(buf, true, M.config.float_opts)
    vim.wo[win].number = false
    vim.wo[win].relativenumber = false
    
    -- Close on q or Esc
    vim.keymap.set('n', 'q', '<cmd>close<cr>', { buffer = buf, nowait = true })
    vim.keymap.set('n', '<Esc>', '<cmd>close<cr>', { buffer = buf, nowait = true })
  end)
end

-- Save CSV back to Excel
function M.save()
  if not M.state.temp_csv or not M.state.original_file then
    vim.notify('No Excel file is currently open', vim.log.levels.WARN)
    return
  end
  
  -- Save the CSV buffer first
  if M.state.excel_buf and vim.api.nvim_buf_is_valid(M.state.excel_buf) then
    vim.api.nvim_buf_call(M.state.excel_buf, function()
      vim.cmd('write')
    end)
  end
  
  local script = string.format([[
import pandas as pd
from openpyxl import load_workbook
import sys

try:
    # Read the modified CSV
    df = pd.read_csv('%s')
    
    # Load existing Excel file
    wb = load_workbook('%s')
    sheet_name = wb.sheetnames[%d]
    ws = wb[sheet_name]
    
    # Clear existing data
    ws.delete_rows(1, ws.max_row)
    
    # Write headers
    for col_idx, col_name in enumerate(df.columns, start=1):
        ws.cell(row=1, column=col_idx, value=col_name)
    
    # Write data
    for row_idx, row in enumerate(df.values, start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    wb.save('%s')
    print('SUCCESS')
except Exception as e:
    print(f'ERROR: {str(e)}', file=sys.stderr)
    sys.exit(1)
]], M.state.temp_csv, M.state.original_file, M.state.current_sheet, M.state.original_file)
  
  run_python(script, function(code, stdout, stderr)
    if code == 0 then
      vim.notify('Excel file saved successfully', vim.log.levels.INFO)
    else
      vim.notify('Error saving Excel file: ' .. stderr, vim.log.levels.ERROR)
    end
  end)
end

-- List sheets
function M.list_sheets()
  if not M.state.original_file then
    vim.notify('No Excel file is currently open', vim.log.levels.WARN)
    return
  end
  
  local lines = { 'Sheets in ' .. vim.fn.fnamemodify(M.state.original_file, ':t') .. ':', '' }
  for i, sheet in ipairs(M.state.sheets) do
    local current = (i - 1 == M.state.current_sheet) and '* ' or '  '
    table.insert(lines, string.format('%s%d. %s', current, i - 1, sheet))
  end
  
  vim.notify(table.concat(lines, '\n'), vim.log.levels.INFO)
end

-- Switch sheet
function M.switch_sheet(sheet_index)
  if not M.state.original_file then
    vim.notify('No Excel file is currently open', vim.log.levels.WARN)
    return
  end
  
  sheet_index = tonumber(sheet_index)
  if not sheet_index or sheet_index < 0 or sheet_index >= #M.state.sheets then
    vim.notify('Invalid sheet index. Use :ExcelSheets to see available sheets', vim.log.levels.ERROR)
    return
  end
  
  M.state.current_sheet = sheet_index
  M.to_csv(M.state.original_file, sheet_index, function(csv_file, err)
    if err then
      vim.notify('Error switching sheet: ' .. err, vim.log.levels.ERROR)
      return
    end
    
    M.state.temp_csv = csv_file
    if M.state.excel_buf and vim.api.nvim_buf_is_valid(M.state.excel_buf) then
      local lines = vim.fn.readfile(csv_file)
      vim.api.nvim_buf_set_lines(M.state.excel_buf, 0, -1, false, lines)
      
      local sheet_name = M.state.sheets[sheet_index + 1]
      vim.api.nvim_buf_set_name(M.state.excel_buf, 
        M.state.original_file .. ' [' .. sheet_name .. ']')
      
      vim.notify('Switched to: ' .. sheet_name, vim.log.levels.INFO)
    end
  end)
end

-- Create new Excel file
function M.create(file)
  local excel_file = get_absolute_path(file or 'new_workbook.xlsx')
  
  local script = string.format([[
from openpyxl import Workbook
import sys

try:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws['A1'] = 'Column1'
    ws['B1'] = 'Column2'
    ws['C1'] = 'Column3'
    wb.save('%s')
    print('SUCCESS')
except Exception as e:
    print(f'ERROR: {str(e)}', file=sys.stderr)
    sys.exit(1)
]], excel_file)
  
  run_python(script, function(code, stdout, stderr)
    if code == 0 then
      vim.notify('Created new Excel file: ' .. excel_file, vim.log.levels.INFO)
      M.open(excel_file)
    else
      vim.notify('Error creating Excel file: ' .. stderr, vim.log.levels.ERROR)
    end
  end)
end

-- Show file info
function M.info()
  if not M.state.original_file then
    vim.notify('No Excel file is currently open', vim.log.levels.WARN)
    return
  end
  
  local info_lines = {
    'Excel File Information:',
    '',
    'File: ' .. M.state.original_file,
    'Current Sheet: ' .. (M.state.sheets[M.state.current_sheet + 1] or 'Unknown'),
    'Total Sheets: ' .. #M.state.sheets,
    'Temp CSV: ' .. (M.state.temp_csv or 'None'),
  }
  
  vim.notify(table.concat(info_lines, '\n'), vim.log.levels.INFO)
end

return M
