-- lua/excel.lua - Core Excel editing functionality (NO-FRICTION VERSION)

local M = {}

--------------------------------------------------
-- State
--------------------------------------------------
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

--------------------------------------------------
-- Config
--------------------------------------------------
M.config = {
	python_cmd = "python3",
	max_col_width = 20,
	min_col_width = 8,
	show_gridlines = true,
	auto_recalc = true,
}

--------------------------------------------------
-- Python handler
--------------------------------------------------
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
			vim.fn.system("cp " .. vim.fn.shellescape(src) .. " " .. vim.fn.shellescape(dst))
		end
	end
	return dst
end

local function call_python(action, args)
	local script = init_python_handler()
	local cmd = { M.config.python_cmd, script, action }

	for _, a in ipairs(args or {}) do
		table.insert(cmd, vim.fn.shellescape(tostring(a)))
	end

	local out = vim.fn.system(table.concat(cmd, " "))
	if vim.v.shell_error ~= 0 then
		vim.notify(out, vim.log.levels.ERROR)
		return nil
	end

	local ok, decoded = pcall(vim.fn.json_decode, out)
	return ok and decoded or out
end

--------------------------------------------------
-- Helpers
--------------------------------------------------
local function col_to_letter(col)
	local s = ""
	while col > 0 do
		local r = (col - 1) % 26
		s = string.char(65 + r) .. s
		col = math.floor((col - 1) / 26)
	end
	return s
end

local function letter_to_col(l)
	local c = 0
	for i = 1, #l do
		c = c * 26 + (string.byte(l, i) - 64)
	end
	return c
end

local function format_value(v, w)
	if not v or v == "" then
		return string.rep(" ", w)
	end
	v = tostring(v)
	if #v > w then
		return v:sub(1, w - 3) .. "..."
	end
	return v .. string.rep(" ", w - #v)
end

--------------------------------------------------
-- Render
--------------------------------------------------
local function render_excel()
	local sheet = M.state.data[M.state.current_sheet]
	if not sheet then
		return {}
	end

	local lines = {}
	local max_row, max_col = 50, 26

	for k in pairs(sheet) do
		local r, c = k:match("(%d+),(%d+)")
		if r and c then
			max_row = math.max(max_row, tonumber(r))
			max_col = math.max(max_col, tonumber(c))
		end
	end

	local header = "   |"
	for c = 1, max_col do
		header = header .. " " .. format_value(col_to_letter(c), M.config.max_col_width) .. " |"
	end
	table.insert(lines, header)

	local sep = "---+"
	for _ = 1, max_col do
		sep = sep .. string.rep("-", M.config.max_col_width + 2) .. "+"
	end
	table.insert(lines, sep)

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

--------------------------------------------------
-- Cursor → Cell
--------------------------------------------------
local function get_cursor_cell()
	local cur = vim.api.nvim_win_get_cursor(0)
	if cur[1] <= 2 or cur[2] < 4 then
		return nil, nil
	end

	local row = cur[1] - 2
	local offset = cur[2] - 4
	local col = math.floor(offset / (M.config.max_col_width + 3)) + 1

	return row, col
end

--------------------------------------------------
-- Buffer update (CRITICAL FIX)
--------------------------------------------------
local function update_buffer()
	if not vim.api.nvim_buf_is_valid(M.state.buffer) then
		return
	end

	local lines = render_excel()
	vim.bo[M.state.buffer].modifiable = true
	vim.api.nvim_buf_set_lines(M.state.buffer, 0, -1, false, lines)
	vim.bo[M.state.buffer].modifiable = false
	vim.bo[M.state.buffer].modified = M.state.modified

	local r, c = M.state.cursor.row, M.state.cursor.col
	if r and c then
		local line = r + 2
		local col = 4 + (c - 1) * (M.config.max_col_width + 3) + 2 -- INSIDE cell value

		if line <= #lines then
			vim.api.nvim_win_set_cursor(0, { line, col })
		end
	end
end

--------------------------------------------------
-- Editing
--------------------------------------------------
function M.edit_cell()
	local r, c = get_cursor_cell()
	if not r or not c then
		return
	end

	local key = r .. "," .. c
	local sheet = M.state.data[M.state.current_sheet] or {}
	local cur = tostring(sheet[key] or "")

	vim.ui.input({
		prompt = col_to_letter(c) .. r .. ": ",
		default = cur,
	}, function(input)
		if input ~= nil then
			M.state.data[M.state.current_sheet] = sheet
			sheet[key] = input
			M.state.modified = true
			M.state.cursor.row = r + 1
			update_buffer()
		end
	end)
end

--------------------------------------------------
-- File ops
--------------------------------------------------
function M.open_excel(path)
	path = path or vim.fn.expand("%:p")
	local res = call_python("load", { path })
	if not res then
		return
	end

	M.state.current_file = path
	M.state.sheets = res.sheets
	M.state.data = res.data
	M.state.current_sheet = res.current_sheet
	M.state.modified = false
	M.state.cursor = { row = 1, col = 1 }

	M.state.buffer = vim.api.nvim_create_buf(false, true)
	vim.api.nvim_set_current_buf(M.state.buffer)
	vim.api.nvim_buf_set_name(M.state.buffer, "Excel")

	-- BUFFER-LOCAL UX FIXES
	vim.bo[M.state.buffer].buftype = "acwrite"
	vim.bo[M.state.buffer].filetype = "excel"
	vim.bo[M.state.buffer].swapfile = false
	vim.bo[M.state.buffer].number = false
	vim.bo[M.state.buffer].relativenumber = false
	vim.bo[M.state.buffer].virtualedit = "onemore"
	vim.bo[M.state.buffer].conceallevel = 0
	vim.bo[M.state.buffer].concealcursor = ""
	vim.bo[M.state.buffer].cursorline = true

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

--------------------------------------------------
-- Keymaps
--------------------------------------------------
function M.setup_keymaps()
	local o = { buffer = M.state.buffer, silent = true }

	vim.keymap.set("n", "<CR>", M.edit_cell, o)
	vim.keymap.set("n", "i", M.edit_cell, o)
	vim.keymap.set("n", "a", M.edit_cell, o)

	vim.keymap.set("n", "h", "h", o)
	vim.keymap.set("n", "j", "j", o)
	vim.keymap.set("n", "k", "k", o)
	vim.keymap.set("n", "l", "l", o)

	vim.keymap.set("n", "<leader>w", M.save_excel, o)
end

--------------------------------------------------
-- Setup
--------------------------------------------------
function M.setup(opts)
	M.config = vim.tbl_deep_extend("force", M.config, opts or {})
end

return M
