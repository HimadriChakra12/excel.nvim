-- lua/excel.lua - Core Excel editing functionality

local M = {}

-- ============================================================================
-- State
-- ============================================================================
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

-- ============================================================================
-- Config
-- ============================================================================
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

-- ============================================================================
-- Python handler
-- ============================================================================
local function get_python_script()
	if M.state.python_script then
		return M.state.python_script
	end
	local path = vim.fn.stdpath("data") .. "/excel_nvim/excel_handler.py"
	M.state.python_script = path
	return path
end

local function init_python_handler()
	local dir = vim.fn.stdpath("data") .. "/excel_nvim"
	vim.fn.mkdir(dir, "p")

	local dst = get_python_script()
	if vim.fn.filereadable(dst) == 0 then
		local plugin_dir = vim.fn.fnamemodify(debug.getinfo(1).source:sub(2), ":h:h")
		local src = plugin_dir .. "/python/excel_handler.py"
		if vim.fn.filereadable(src) == 1 then
			vim.fn.system({
				"cp",
				vim.fn.shellescape(src),
				vim.fn.shellescape(dst),
			})
		end
	end

	return dst
end

local function call_python(action, args)
	local script = init_python_handler()
	local cmd = { M.config.python_cmd, script, action }

	for _, arg in ipairs(args or {}) do
		table.insert(cmd, vim.fn.shellescape(tostring(arg)))
	end

	local out = vim.fn.system(table.concat(cmd, " "))
	if vim.v.shell_error ~= 0 then
		vim.notify("Excel error: " .. out, vim.log.levels.ERROR)
		return nil
	end

	local ok, decoded = pcall(vim.fn.json_decode, out)
	return ok and decoded or out
end

-- ============================================================================
-- Helpers
-- ============================================================================
local function col_to_letter(col)
	local s = ""
	while col > 0 do
		local r = (col - 1) % 26
		s = string.char(65 + r) .. s
		col = math.floor((col - 1) / 26)
	end
	return s
end

local function letter_to_col(letters)
	local col = 0
	for i = 1, #letters do
		col = col * 26 + (letters:byte(i) - 64)
	end
	return col
end

local function format_value(value, width)
	if not value or value == "" then
		return string.rep(" ", width)
	end
	local s = tostring(value)
	if #s > width then
		return s:sub(1, width - 3) .. "..."
	end
	return s .. string.rep(" ", width - #s)
end

-- ============================================================================
-- Rendering
-- ============================================================================
local function render_excel()
	local sheet = M.state.data[M.state.current_sheet]
	if not sheet then
		return {}
	end

	local max_row, max_col = 50, 26
	for k in pairs(sheet) do
		local r, c = k:match("(%d+),(%d+)")
		if r and c then
			max_row = math.max(max_row, tonumber(r))
			max_col = math.max(max_col, tonumber(c))
		end
	end

	local lines = {}

	-- Header
	local header = "   |"
	for c = 1, max_col do
		header = header .. " " .. format_value(col_to_letter(c), M.config.max_col_width) .. " |"
	end
	table.insert(lines, header)

	-- Separator
	local sep = "---+"
	for _ = 1, max_col do
		sep = sep .. string.rep("-", M.config.max_col_width + 2) .. "+"
	end
	table.insert(lines, sep)

	-- Data
	for r = 1, max_row do
		local line = string.format("%3d|", r)
		for c = 1, max_col do
			local v = sheet[r .. "," .. c] or ""
			line = line .. " " .. format_value(v, M.config.max_col_width) .. " |"
		end
		table.insert(lines, line)
	end

	return lines
end

local function get_cursor_cell()
	local pos = vim.api.nvim_win_get_cursor(0)
	local line, col = pos[1], pos[2]

	if line <= 2 or col < 4 then
		return nil, nil
	end

	local row = line - 2
	local offset = col - 4
	local cell_col = math.floor(offset / (M.config.max_col_width + 3)) + 1

	return row, cell_col
end

local function update_buffer()
	if not M.state.buffer or not vim.api.nvim_buf_is_valid(M.state.buffer) then
		return
	end

	local lines = render_excel()
	vim.api.nvim_buf_set_lines(M.state.buffer, 0, -1, false, lines)
	vim.api.nvim_buf_set_option(M.state.buffer, "modified", M.state.modified)

	if M.state.cursor then
		local l = M.state.cursor.row + 2
		local c = 4 + (M.state.cursor.col - 1) * (M.config.max_col_width + 3) + 1
		if l <= #lines then
			vim.api.nvim_win_set_cursor(0, { l, c })
		end
	end
end

-- ============================================================================
-- Public API
-- ============================================================================
function M.open_excel(filepath)
	filepath = filepath or vim.fn.expand("%:p")
	if filepath == "" then
		vim.notify("No file specified", vim.log.levels.ERROR)
		return
	end

	local res = call_python("load", { filepath })
	if not res then
		return
	end

	M.state.current_file = filepath
	M.state.sheets = res.sheets or {}
	M.state.data = res.data or {}
	M.state.current_sheet = res.current_sheet or M.state.sheets[1]
	M.state.modified = false
	M.state.cursor = { row = 1, col = 1 }

	M.state.buffer = vim.api.nvim_create_buf(false, true)
	vim.api.nvim_buf_set_name(M.state.buffer, "Excel: " .. vim.fn.fnamemodify(filepath, ":t"))

	vim.api.nvim_buf_set_option(M.state.buffer, "buftype", "nofile")
	vim.api.nvim_buf_set_option(M.state.buffer, "swapfile", false)
	vim.api.nvim_buf_set_option(M.state.buffer, "modifiable", true)
	vim.api.nvim_buf_set_option(M.state.buffer, "filetype", "excel")

	vim.api.nvim_set_current_buf(M.state.buffer)
	M.setup_keymaps()
	update_buffer()
end

function M.save_excel()
	if not M.state.current_file then
		return
	end
	call_python("save", {
		M.state.current_file,
		vim.fn.json_encode(M.state.data),
		M.state.current_sheet,
	})
	M.state.modified = false
end

function M.edit_cell()
	local row, col = get_cursor_cell()
	if not row or not col then
		return
	end

	local key = row .. "," .. col
	local sheet = M.state.data[M.state.current_sheet]
	local current = sheet[key] or ""

	vim.ui.input({
		prompt = col_to_letter(col) .. row .. ": ",
		default = tostring(current),
	}, function(input)
		if input ~= nil then
			sheet[key] = input
			M.state.modified = true
			M.state.cursor = { row = row + 1, col = col }
			update_buffer()
		end
	end)
end

-- ============================================================================
-- Keymaps
-- ============================================================================
function M.setup_keymaps()
	local opts = { buffer = M.state.buffer, silent = true }

	vim.keymap.set("n", "<CR>", M.edit_cell, opts)
	vim.keymap.set("n", "i", M.edit_cell, opts)
	vim.keymap.set("n", "a", M.edit_cell, opts)
	vim.keymap.set("n", "<leader>w", M.save_excel, opts)
end

return M
