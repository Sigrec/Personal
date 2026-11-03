-- Git status signs next to file names (needs the git fetchers in yazi.toml)
require("git"):setup {
	order = 1500,
}

-- Border around the panes
require("full-border"):setup {
	set = "rounded",
}
