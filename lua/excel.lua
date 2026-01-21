-- lua/excel.lua - Core Excel editing functionality

local M = {}

-- State management
M.state = {
	current_file = nil,
	current_sheet = nil,
	sheets = {},
	data = {},
	modified = false,
	cursor = { row = 1, col = 1 },
	buffer = nil,
	namespace = vim.api.nvim_create_namespace("excel_nvim"),
	python_script = nil,
}

-- Configuration
M.config = {
	python_cmd = "python3",
	max_col_width = 20,
	min_col_width = 8,
	show_gridlines = true,
	show_formulas = false,
	auto_recalc = true,
	date_format = "%Y-%m-%d",
	number_format = "%.2f",
	cell_padding = 1,
}

-- Get the Python script path
local function get_python_script()
	if M.state.python_script then
		return M.state.python_script
	end

	local script_path = vim.fn.stdpath("data") .. "/excel_nvim/excel_handler.py"
	M.state.python_script = script_path
	return script_path
end

-- Initialize Python handler
local function init_python_handler()
	local script_dir = vim.fn.stdpath("data") .. "/excel_nvim"
	vim.fn.mkdir(script_dir, "p")

	local script_path = get_python_script()

	-- Copy the Python script if it doesn't exist
	if vim.fn.filereadable(script_path) == 0 then
		local plugin_dir = vim.fn.fnamemodify(debug.getinfo(1).source:sub(2), ":h:h")
		local source_script = plugin_dir .. "/python/excel_handler.py"

		if vim.fn.filereadable(source_script) == 1 then
			vim.fn.system("cp " .. vim.fn.shellescape(source_script) .. " " .. vim.fn.shellescape(script_path))
		end
	end

	return script_path
end

-- Call Python script
local function call_python(action, args)
	local script_path = init_python_handler()

	local cmd = {
		M.config.python_cmd,
		script_path,
		action,
	}

	for _, arg in ipairs(args or {}) do
		table.insert(cmd, vim.fn.shellescape(tostring(arg)))
	end

	local result = vim.fn.system(table.concat(cmd, " "))

	if vim.v.shell_error ~= 0 then
		vim.notify("Excel operation failed: " .. result, vim.log.levels.ERROR)
		return nil
	end

	local ok, decoded = pcall(vim.fn.json_decode, result)
	if not ok then
		return result
	end

	return decoded
end

-- Convert column index to letter (1 -> A, 27 -> AA)
local function col_to_letter(col)
	local letter = ""
	while col > 0 do
		local rem = (col - 1) % 26
		letter = string.char(65 + rem) .. letter
		col = math.floor((col - 1) / 26)
	end
	return letter
end

-- Convert column letter to index (A -> 1, AA -> 27)
local function letter_to_col(letter)
	local col = 0
	for i = 1, #letter do
		col = col * 26 + (string.byte(letter, i) - 64)
	end
	return col
end

-- Format cell value for display
local function format_value(value, width)
	if value == nil or value == "" then
		return string.rep(" ", width)
	end

	local str = tostring(value)
	if #str > width then
		return str:sub(1, width - 3) .. "..."
	end

	return str .. string.rep(" ", width - #str)
end

-- Render Excel data as text
local function render_excel()
	if not M.state.data or not M.state.current_sheet then
		return {}
	end

	local sheet_data = M.state.data[M.state.current_sheet]
	if not sheet_data then
		return {}
	end

	local lines = {}
	local max_row = 0
	local max_col = 0

	-- Find dimensions
	for cell, _ in pairs(sheet_data) do
		local row, col = cell:match("(%d+),(%d+)")
		if row and col then
			row, col = tonumber(row), tonumber(col)
			max_row = math.max(max_row, row)
			max_col = math.max(max_col, col)
		end
	end

	-- Ensure minimum size
	max_row = math.max(max_row, 50)
	max_col = math.max(max_col, 26)

	-- Header row with column letters
	local header = "   |"
	for col = 1, max_col do
		local col_letter = col_to_letter(col)
		header = header .. " " .. format_value(col_letter, M.config.max_col_width) .. " |"
	end
	table.insert(lines, header)

	-- Separator
	local sep = "---+"
	for col = 1, max_col do
		sep = sep .. string.rep("-", M.config.max_col_width + 2) .. "+"
	end
	table.insert(lines, sep)

	-- Data rows
	for row = 1, max_row do
		local line = string.format("%3d|", row)
		for col = 1, max_col do
			local cell_key = row .. "," .. col
			local value = sheet_data[cell_key] or ""
			line = line .. " " .. format_value(value, M.config.max_col_width) .. " |"
		end
		table.insert(lines, line)
	end

	return lines
end

-- Get cell at cursor
local function get_cursor_cell()
	local cursor = vim.api.nvim_win_get_cursor(0)
	local line = cursor[1]
	local col = cursor[2]

	-- Line 1: Header row (column letters)
	-- Line 2: Separator (---)
	-- Line 3+: Data rows (row 1, 2, 3...)

	if line <= 2 then -- Header or separator
		return nil, nil
	end

	-- Calculate data row: Line 3 = Row 1, Line 4 = Row 2, etc.
	local row = line - 2

	-- Column calculation:
	-- Format: "NNN|" (4 chars for row number) then " VALUE | VALUE |..."
	-- Each cell is: " " + value (max_col_width chars) + " |" = max_col_width + 3

	if col < 4 then
		-- Cursor is on the row number area
		return nil, nil
	end

	-- Subtract the row number area (4 chars)
	local offset = col - 4

	-- Each column takes up (max_col_width + 3) characters
	local cell_col = math.floor(offset / (M.config.max_col_width + 3)) + 1

	return row, cell_col
end

-- Update buffer with Excel data
local function update_buffer()
	if not M.state.buffer or not vim.api.nvim_buf_is_valid(M.state.buffer) then
		return
	end

	local lines = render_excel()

	vim.api.nvim_buf_set_option(M.state.buffer, "modifiable", true)
	vim.api.nvim_buf_set_lines(M.state.buffer, 0, -1, false, lines)
	vim.api.nvim_buf_set_option(M.state.buffer, "modifiable", false)
	vim.api.nvim_buf_set_option(M.state.buffer, "modified", M.state.modified)

	-- Set cursor position
	if M.state.cursor.row and M.state.cursor.col then
		local line = M.state.cursor.row + 2 -- +2 for header and separator (display line)
		-- Calculate horizontal position: 4 (row number area) + (col-1) * (width + 3) + 1 (into cell)
		local col = 4 + (M.state.cursor.col - 1) * (M.config.max_col_width + 3) + 1

		-- Ensure we're within bounds
		if line <= #lines then
			vim.api.nvim_win_set_cursor(0, { line, col })
		end
	end

	-- Update status line with current cell info
	local row, col = get_cursor_cell()
	if row and col then
		local cell_ref = col_to_letter(col) .. row
		vim.b.excel_cell = cell_ref
	end
end

-- Open Excel file
function M.open_excel(filepath)
	filepath = filepath or vim.fn.expand("%:p")

	if filepath == "" then
		vim.notify("No file specified", vim.log.levels.ERROR)
		return
	end

	vim.notify("Loading Excel file: " .. filepath, vim.log.levels.INFO)

	local result = call_python("load", { filepath })

	if not result then
		return
	end

	M.state.current_file = filepath
	M.state.sheets = result.sheets or {}
	M.state.data = result.data or {}
	M.state.current_sheet = result.current_sheet or (M.state.sheets[1] or "Sheet1")
	M.state.modified = false

	-- Create buffer
	M.state.buffer = vim.api.nvim_create_buf(false, true)
	vim.api.nvim_buf_set_name(M.state.buffer, "Excel: " .. vim.fn.fnamemodify(filepath, ":t"))

	-- Set buffer options
	vim.api.nvim_buf_set_option(M.state.buffer, "buftype", "acwrite")
	vim.api.nvim_buf_set_option(M.state.buffer, "filetype", "excel")
	vim.api.nvim_buf_set_option(M.state.buffer, "swapfile", false)

	-- Switch to buffer
	vim.api.nvim_set_current_buf(M.state.buffer)

	-- Set up keymaps
	M.setup_keymaps()

	-- Render
	update_buffer()

	vim.notify("Loaded sheet: " .. M.state.current_sheet .. " (" .. #M.state.sheets .. " sheets)", vim.log.levels.INFO)
end

-- Save Excel file
function M.save_excel(filepath)
	filepath = filepath or M.state.current_file

	if not filepath then
		vim.notify("No file to save", vim.log.levels.ERROR)
		return
	end

	vim.notify("Saving Excel file: " .. filepath, vim.log.levels.INFO)

	local result = call_python("save", {
		filepath,
		vim.fn.json_encode(M.state.data),
		M.state.current_sheet,
	})

	if result and result.success then
		M.state.modified = false
		vim.api.nvim_buf_set_option(M.state.buffer, "modified", false)
		vim.notify("Saved: " .. filepath, vim.log.levels.INFO)
	end
end

-- Edit cell
function M.edit_cell()
	local row, col = get_cursor_cell()

	if not row or not col then
		local cursor = vim.api.nvim_win_get_cursor(0)
		vim.notify(string.format("Not on a valid cell (line=%d, col=%d)", cursor[1], cursor[2]), vim.log.levels.WARN)
		return
	end

	local cell_ref = col_to_letter(col) .. row
	local cell_key = row .. "," .. col
	local current_value = ""

	if M.state.data[M.state.current_sheet] then
		current_value = M.state.data[M.state.current_sheet][cell_key] or ""
	end

	-- Convert to string if it's not already
	current_value = tostring(current_value)

	-- Debug info
	local cursor = vim.api.nvim_win_get_cursor(0)
	local debug_msg =
		string.format("[DEBUG] Line=%d Col=%d -> Row=%d CellCol=%d (%s)", cursor[1], cursor[2], row, col, cell_ref)
	vim.notify(debug_msg, vim.log.levels.INFO)

	vim.ui.input({
		prompt = string.format("Cell %s [%s]: ", cell_ref, M.state.current_sheet),
		default = current_value,
	}, function(input)
		if input ~= nil then
			if not M.state.data[M.state.current_sheet] then
				M.state.data[M.state.current_sheet] = {}
			end

			-- Store the value (empty string if user cleared it)
			M.state.data[M.state.current_sheet][cell_key] = input
			M.state.modified = true
			update_buffer()

			-- Move cursor down after editing (Excel-like behavior)
			if M.state.cursor.row then
				M.state.cursor.row = M.state.cursor.row + 1
				update_buffer()
			end
		end
	end)
end

-- Insert formula
function M.insert_formula(formula)
	local row, col = get_cursor_cell()

	if not row or not col then
		vim.notify("Not on a valid cell", vim.log.levels.WARN)
		return
	end

	if not formula or formula == "" then
		vim.ui.input({
			prompt = "Enter formula (e.g., =SUM(A1:A10)): ",
		}, function(input)
			if input and input ~= "" then
				M.insert_formula(input)
			end
		end)
		return
	end

	local cell_key = row .. "," .. col

	if not M.state.data[M.state.current_sheet] then
		M.state.data[M.state.current_sheet] = {}
	end

	M.state.data[M.state.current_sheet][cell_key] = formula
	M.state.modified = true
	update_buffer()

	vim.notify("Formula inserted: " .. formula, vim.log.levels.INFO)
end

-- List sheets
function M.list_sheets()
	if #M.state.sheets == 0 then
		vim.notify("No sheets available", vim.log.levels.WARN)
		return
	end

	local items = {}
	for i, sheet in ipairs(M.state.sheets) do
		local marker = sheet == M.state.current_sheet and "* " or "  "
		table.insert(items, string.format("%s%d. %s", marker, i, sheet))
	end

	vim.ui.select(items, {
		prompt = "Select sheet:",
	}, function(choice, idx)
		if idx then
			M.switch_sheet(M.state.sheets[idx])
		end
	end)
end

-- Switch sheet
function M.switch_sheet(sheet_name)
	if not sheet_name or sheet_name == "" then
		M.list_sheets()
		return
	end

	local found = false
	for _, sheet in ipairs(M.state.sheets) do
		if sheet == sheet_name then
			found = true
			break
		end
	end

	if not found then
		vim.notify("Sheet not found: " .. sheet_name, vim.log.levels.ERROR)
		return
	end

	M.state.current_sheet = sheet_name
	M.state.cursor = { row = 1, col = 1 }
	update_buffer()

	vim.notify("Switched to sheet: " .. sheet_name, vim.log.levels.INFO)
end

-- New sheet
function M.new_sheet(name)
	if not name or name == "" then
		vim.ui.input({
			prompt = "Enter sheet name: ",
			default = "Sheet" .. (#M.state.sheets + 1),
		}, function(input)
			if input and input ~= "" then
				M.new_sheet(input)
			end
		end)
		return
	end

	for _, sheet in ipairs(M.state.sheets) do
		if sheet == name then
			vim.notify("Sheet already exists: " .. name, vim.log.levels.ERROR)
			return
		end
	end

	table.insert(M.state.sheets, name)
	M.state.data[name] = {}
	M.state.current_sheet = name
	M.state.modified = true
	update_buffer()

	vim.notify("Created new sheet: " .. name, vim.log.levels.INFO)
end

-- Delete sheet
function M.delete_sheet(name)
	name = name or M.state.current_sheet

	if #M.state.sheets <= 1 then
		vim.notify("Cannot delete the last sheet", vim.log.levels.ERROR)
		return
	end

	for i, sheet in ipairs(M.state.sheets) do
		if sheet == name then
			table.remove(M.state.sheets, i)
			M.state.data[name] = nil

			if M.state.current_sheet == name then
				M.state.current_sheet = M.state.sheets[1]
			end

			M.state.modified = true
			update_buffer()
			vim.notify("Deleted sheet: " .. name, vim.log.levels.INFO)
			return
		end
	end

	vim.notify("Sheet not found: " .. name, vim.log.levels.ERROR)
end

-- Insert row
function M.insert_row(position)
	local row, _ = get_cursor_cell()
	position = position and tonumber(position) or row

	if not position then
		vim.notify("Invalid row position", vim.log.levels.ERROR)
		return
	end

	local sheet_data = M.state.data[M.state.current_sheet] or {}
	local new_data = {}

	for cell_key, value in pairs(sheet_data) do
		local r, c = cell_key:match("(%d+),(%d+)")
		r, c = tonumber(r), tonumber(c)

		if r >= position then
			new_data[(r + 1) .. "," .. c] = value
		else
			new_data[cell_key] = value
		end
	end

	M.state.data[M.state.current_sheet] = new_data
	M.state.modified = true
	update_buffer()

	vim.notify("Inserted row at position " .. position, vim.log.levels.INFO)
end

-- Insert column
function M.insert_column(position)
	local _, col = get_cursor_cell()
	position = position and tonumber(position) or col

	if not position then
		vim.notify("Invalid column position", vim.log.levels.ERROR)
		return
	end

	local sheet_data = M.state.data[M.state.current_sheet] or {}
	local new_data = {}

	for cell_key, value in pairs(sheet_data) do
		local r, c = cell_key:match("(%d+),(%d+)")
		r, c = tonumber(r), tonumber(c)

		if c >= position then
			new_data[r .. "," .. (c + 1)] = value
		else
			new_data[cell_key] = value
		end
	end

	M.state.data[M.state.current_sheet] = new_data
	M.state.modified = true
	update_buffer()

	vim.notify("Inserted column at position " .. position, vim.log.levels.INFO)
end

-- Delete row
function M.delete_row(position)
	local row, _ = get_cursor_cell()
	position = position and tonumber(position) or row

	if not position then
		vim.notify("Invalid row position", vim.log.levels.ERROR)
		return
	end

	local sheet_data = M.state.data[M.state.current_sheet] or {}
	local new_data = {}

	for cell_key, value in pairs(sheet_data) do
		local r, c = cell_key:match("(%d+),(%d+)")
		r, c = tonumber(r), tonumber(c)

		if r < position then
			new_data[cell_key] = value
		elseif r > position then
			new_data[(r - 1) .. "," .. c] = value
		end
	end

	M.state.data[M.state.current_sheet] = new_data
	M.state.modified = true
	update_buffer()

	vim.notify("Deleted row " .. position, vim.log.levels.INFO)
end

-- Delete column
function M.delete_column(position)
	local _, col = get_cursor_cell()
	position = position and tonumber(position) or col

	if not position then
		vim.notify("Invalid column position", vim.log.levels.ERROR)
		return
	end

	local sheet_data = M.state.data[M.state.current_sheet] or {}
	local new_data = {}

	for cell_key, value in pairs(sheet_data) do
		local r, c = cell_key:match("(%d+),(%d+)")
		r, c = tonumber(r), tonumber(c)

		if c < position then
			new_data[cell_key] = value
		elseif c > position then
			new_data[r .. "," .. (c - 1)] = value
		end
	end

	M.state.data[M.state.current_sheet] = new_data
	M.state.modified = true
	update_buffer()

	vim.notify("Deleted column " .. position, vim.log.levels.INFO)
end

-- Go to cell
function M.goto_cell(cell_ref)
	if not cell_ref or cell_ref == "" then
		vim.ui.input({
			prompt = "Go to cell (e.g., B5): ",
		}, function(input)
			if input and input ~= "" then
				M.goto_cell(input)
			end
		end)
		return
	end

	local col_part, row_part = cell_ref:match("([A-Z]+)(%d+)")

	if not col_part or not row_part then
		vim.notify("Invalid cell reference: " .. cell_ref, vim.log.levels.ERROR)
		return
	end

	local col = letter_to_col(col_part)
	local row = tonumber(row_part)

	M.state.cursor = { row = row, col = col }
	update_buffer()
end

-- Format cell
function M.format_cell()
	local row, col = get_cursor_cell()

	if not row or not col then
		vim.notify("Not on a valid cell", vim.log.levels.WARN)
		return
	end

	vim.notify("Cell formatting coming in future update", vim.log.levels.INFO)
end

-- Recalculate formulas
function M.recalculate()
	if not M.state.current_file then
		vim.notify("No file loaded", vim.log.levels.ERROR)
		return
	end

	vim.notify("Recalculating formulas...", vim.log.levels.INFO)

	-- Save first
	M.save_excel()

	-- Run recalculation
	local result = call_python("recalc", { M.state.current_file })

	if result and result.success then
		vim.notify("Formulas recalculated successfully", vim.log.levels.INFO)
		-- Reload the file
		M.open_excel(M.state.current_file)
	end
end

-- Freeze panes
function M.freeze_panes(position)
	vim.notify("Freeze panes feature coming in future update", vim.log.levels.INFO)
end

-- Sort range
function M.sort_range(options)
	vim.notify("Sort feature coming in future update", vim.log.levels.INFO)
end

-- Toggle filter
function M.toggle_filter()
	vim.notify("Filter feature coming in future update", vim.log.levels.INFO)
end

-- Create chart
function M.create_chart(...)
	vim.notify("Chart feature coming in future update", vim.log.levels.INFO)
end

-- Set up keymaps
function M.setup_keymaps()
	local opts = { buffer = M.state.buffer, silent = true }

	-- Navigation
	vim.keymap.set("n", "<CR>", M.edit_cell, opts)
	vim.keymap.set("n", "i", M.edit_cell, opts)
	vim.keymap.set("n", "a", M.edit_cell, opts)
	vim.keymap.set("n", "ge", M.goto_cell, opts)

	-- Update cursor state on movement
	local function update_cursor_state()
		local row, col = get_cursor_cell()
		if row and col then
			M.state.cursor = { row = row, col = col }
			-- Update status line
			local cell_ref = col_to_letter(col) .. row
			vim.b.excel_cell = cell_ref
		end
	end

	-- Track cursor movement
	vim.keymap.set("n", "h", function()
		vim.cmd("normal! h")
		update_cursor_state()
	end, opts)

	vim.keymap.set("n", "j", function()
		vim.cmd("normal! j")
		update_cursor_state()
	end, opts)

	vim.keymap.set("n", "k", function()
		vim.cmd("normal! k")
		update_cursor_state()
	end, opts)

	vim.keymap.set("n", "l", function()
		vim.cmd("normal! l")
		update_cursor_state()
	end, opts)

	-- Sheets
	vim.keymap.set("n", "gs", M.list_sheets, opts)
	vim.keymap.set("n", "gn", M.new_sheet, opts)
	vim.keymap.set("n", "gd", M.delete_sheet, opts)

	-- Rows and columns
	vim.keymap.set("n", "ir", M.insert_row, opts)
	vim.keymap.set("n", "ic", M.insert_column, opts)
	vim.keymap.set("n", "dr", M.delete_row, opts)
	vim.keymap.set("n", "dc", M.delete_column, opts)

	-- Formula
	vim.keymap.set("n", "gf", M.insert_formula, opts)
	vim.keymap.set("n", "gr", M.recalculate, opts)

	-- File operations
	vim.keymap.set("n", "<leader>w", M.save_excel, opts)
	vim.keymap.set("n", ":w<CR>", M.save_excel, opts)
end

-- Setup function
function M.setup(opts)
	M.config = vim.tbl_deep_extend("force", M.config, opts or {})
end

return M
