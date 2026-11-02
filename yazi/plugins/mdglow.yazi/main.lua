-- Renders Markdown in the preview pane with glow. Written against the yazi 26.x
-- plugin API (Reledia/glow still uses the removed ya.mgr_emit / ya.preview_widgets).
local M = {}

-- glow's dark style minus the "##" heading markers, kept next to yazi.toml.
-- Falls back to the stock dark style if the config folder is not set.
local function style_path()
	local home = os and os.getenv and os.getenv("YAZI_CONFIG_HOME")
	return home and (home .. "\\glow-style.json") or "dark"
end

function M:peek(job)
	local style = style_path()
	-- No <close> here: the Child object in the 26.9.1 release is not closable
	local child = Command("glow")
		:arg({ "--style", style, "--width", tostring(job.area.w), tostring(job.file.path) })
		:env("CLICOLOR_FORCE", "1")
		:stdout(Command.PIPED)
		:stderr(Command.PIPED)
		:spawn()

	if not child then
		return require("code"):peek(job)
	end

	local opt = { ansi = true, tab_size = rt.preview.tab_size, wrap = rt.preview.wrap, width = job.area.w }
	local limit = job.area.h
	local i, lines = 0, {}
	repeat
		local next, event = child:read_line()
		if event == 1 then
			child:start_kill()
			return require("code"):peek(job)
		elseif event ~= 0 then
			break
		end

		local wrapped = ui.lines(next, opt)
		local from = math.max(1, job.skip - i + 1)
		local to = math.min(#wrapped, job.skip + limit - i)

		i = i + #wrapped
		for j = from, to do
			lines[#lines + 1] = wrapped[j]
		end
	until i >= job.skip + limit

	child:start_kill()
	if job.skip > 0 and i < job.skip + limit then
		ya.emit("peek", { math.max(0, i - limit), only_if = job.file.url, upper_bound = true, scope = job.scope })
	else
		ya.preview_widget(job, ui.Text(lines):area(job.area))
	end
end

function M:seek(job) require("code"):seek(job) end

return M
