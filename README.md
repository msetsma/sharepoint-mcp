# sharepoint-mcp

[![CI](https://github.com/msetsma/sharepoint-mcp/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/msetsma/sharepoint-mcp/actions/workflows/ci.yml)

Read-only-by-default MCP server for SharePoint on macOS and Windows: OneDrive-synced
SharePoint libraries read from disk, and SharePoint lists via Microsoft Graph.
The file tools need no Microsoft Graph consent and no Full Disk Access; the list
tools (`sp_lists`, `sp_list_items`) use the Azure CLI's delegated Graph token
instead, so they require a signed-in `az`. List edits exist too, but are
[opt-in](#writes-opt-in-update-only).

## Tools

| Tool | What it does |
|---|---|
| `sp_roots` | Synced SharePoint library roots |
| `sp_find` | Filename search; downloads nothing |
| `sp_read` | Text of one file; downloads it if needed |
| `sp_lists` | Lists and libraries on a SharePoint site (Graph) |
| `sp_list_items` | Rows with column values from a list (Graph) |
| `sp_my_sites` | Sites behind my Teams and Microsoft 365 groups, with their ids (Graph) |
| `sp_list_schema` | A list's columns: internal names, types, allowed choices, read-only and indexed flags (Graph) |
| `sp_recent_changes` | Files or rows changed since a date, newest first, with who changed them; works on unsynced libraries (Graph) |

The Graph tools borrow the `az` CLI's delegated Graph token, so run `az login`
once before using them (no app registration needed). That token reaches sites
through group membership, so `sp_my_sites` only lists group-connected sites;
pass any other site's URL to the tools directly. Tenants can restrict what the
Azure CLI may read, in which case a tool returns Graph's 403 as an error.

## Writes (opt-in, update only)

> [!WARNING]
> Setting `SHAREPOINT_ENABLE_WRITES` to a truthy value (`1`, `true`, `yes`, `on`)
> registers one tool, `sp_list_item_update`, which changes existing SharePoint
> list items. Without it, the server is read-only.

There is no create or delete tool, by design. Agents can edit rows but never
add or remove them.

Set it in the `env` of the server's client config (see [Run](#run)):

```json
{
  "command": "/absolute/path/to/uv",
  "args": ["run", "--directory", "/absolute/path/to/sharepoint-mcp", "server.py"],
  "env": { "SHAREPOINT_ENABLE_WRITES": "1" }
}
```

| Tool | What it does |
|---|---|
| `sp_list_item_update` | Change only the given fields on an existing item; returns each field's value before and after |

`fields` is a dict of column values, e.g. `{"status": "In Progress"}`. The tool
registers with `read_only_hint=False` and `destructive_hint=True`, so clients
prompt before each call. The `az` sign-in needs write access to the site, or
Graph returns 403.

## Requirements

- macOS or Windows, with OneDrive signed in to a work account (file tools). Libraries added with **Sync** or with **Add shortcut to My files** both work
- [uv](https://docs.astral.sh/uv/) (fetches Python 3.14 on first run)
- A signed-in Azure CLI (`az login`) for the list tools

## Run

```sh
git clone https://github.com/msetsma/sharepoint-mcp
cd sharepoint-mcp
uv run server.py        # stdio MCP server
uv run test_server.py   # offline self-check
```

Register it with an MCP client as a stdio server:

```json
{
  "command": "/absolute/path/to/uv",
  "args": ["run", "--directory", "/absolute/path/to/sharepoint-mcp", "server.py"]
}
```

Use an absolute path to `uv` in GUI apps (find it with `which uv`); they don't inherit your shell `PATH`.

## macOS

### Install with an AI agent

Paste this into Claude Code (or another coding agent) on the Mac:

````text
Install and configure the sharepoint-mcp MCP server (https://github.com/msetsma/sharepoint-mcp) on this Mac. Do each step, check it worked, and stop to tell me if one fails.

1. Install Homebrew packages if missing: `brew install uv azure-cli`. Check with `uv --version` and `az version`.
2. Clone the repo to ~/sharepoint-mcp (skip if it exists; run `git pull` instead).
3. Run `uv run --directory ~/sharepoint-mcp test_server.py`. All tests must pass.
4. Check `az account show`. If I'm not signed in, tell me to run `az login` myself with my work account, and wait for me to confirm.
5. Check that OneDrive is set up: `ls ~/Library/CloudStorage | grep '^OneDrive-'` should list at least one folder other than OneDrive-Personal. The server reads all of them, so libraries added with Sync and ones added with "Add shortcut to My files" both work. If there are none, tell me to sign in to OneDrive with my work account. The list tools still work without this, so continue.
6. Find the absolute path of uv with `which uv`.
7. Register the server with every client installed here, using that absolute uv path:
   - Claude Code: `claude mcp add --scope user sharepoint-mcp -- <uv> run --directory <home>/sharepoint-mcp server.py`
   - Claude Desktop: add an entry under "mcpServers" in ~/Library/Application Support/Claude/claude_desktop_config.json with "command": "<uv>", "args": ["run", "--directory", "<home>/sharepoint-mcp", "server.py"], and "env": {"PATH": "/opt/homebrew/bin:/usr/local/bin:/usr/bin:/bin:/usr/sbin:/sbin"}. Merge into the existing file; don't overwrite other servers. Back up the file first.
   Use real absolute paths, not ~.
8. Ask me whether to enable list writes, and explain first: setting SHAREPOINT_ENABLE_WRITES=1 adds one tool, sp_list_item_update, which edits existing SharePoint list items (no create or delete) as my `az` account, and clients ask before each call. Default to no. Only if I say yes, add "SHAREPOINT_ENABLE_WRITES": "1" to the env of each config entry (for Claude Code, re-run the add command with `-e SHAREPOINT_ENABLE_WRITES=1` before `--`). Never set it without my yes.
9. Verify: `claude mcp list` shows sharepoint-mcp as connected. Tell me to restart Claude Desktop if you edited its config.
10. Summarise what you installed, where the config entries are, and anything I still need to do.
````

## Windows

The server runs on Windows with the same tools. Three things differ from
macOS, and the code handles each one:

| | macOS | Windows |
|---|---|---|
| Where SharePoint files live | every `~/Library/CloudStorage/OneDrive-*` folder except `OneDrive-Personal`: synced libraries (`OneDrive-SharedLibraries-<Org>` or `OneDrive-<Org>-<Site>`) and the work OneDrive (`OneDrive-<Org>`, which holds shortcut-added libraries) | read from OneDrive's registry entries under `HKCU\Software\SyncEngines\Providers\OneDrive`: `%USERPROFILE%\<Org>` for synced libraries and `%USERPROFILE%\OneDrive - <Org>` for the work OneDrive. Consumer OneDrive is skipped |
| How a cloud-only file is spotted | `st_blocks == 0` | `FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS` or `FILE_ATTRIBUTE_OFFLINE` set |
| How `az` is launched | the `az` script | the CLI's bundled `python.exe -IBm azure.cli`, bypassing `az.cmd` so `cmd.exe` never re-parses `&`, `%` or `"` in URLs and JSON bodies |

### Install with an AI agent

Paste this into Claude Code (or another coding agent) on the Windows PC:

````text
Install and configure the sharepoint-mcp MCP server (https://github.com/msetsma/sharepoint-mcp) on this Windows PC. Use PowerShell. Do each step, check it worked, and stop to tell me if one fails.

1. Install if missing: `winget install astral-sh.uv` and `winget install Microsoft.AzureCLI`. Open a new shell so PATH updates, then check `uv --version` and `az version`. The Azure CLI must be the winget/MSI install, not pip.
2. Clone the repo to $HOME\sharepoint-mcp (skip if it exists; run `git pull` instead). Install Git with `winget install Git.Git` if needed.
3. Run `uv run --directory $HOME\sharepoint-mcp test_server.py`. All tests must pass.
4. Check `az account show`. If I'm not signed in, tell me to run `az login` myself with my work account, and wait for me to confirm.
5. Check that OneDrive is set up: `Get-ChildItem HKCU:\Software\SyncEngines\Providers\OneDrive | Get-ItemProperty | Select-Object MountPoint, UrlNamespace`. The server reads every entry whose UrlNamespace contains ".sharepoint.com", so libraries added with Sync and ones added with "Add shortcut to My files" both work. If there are none, tell me to sign in to OneDrive with my work account. The list tools still work without this, so continue.
6. Find the absolute path of uv with `(Get-Command uv).Source`.
7. Register the server with every client installed here, using that absolute uv.exe path:
   - Claude Code: `claude mcp add --scope user sharepoint-mcp -- <uv.exe> run --directory <home>\sharepoint-mcp server.py`
   - Claude Desktop: add an entry under "mcpServers" in its config, with "command": "<uv.exe>" and "args": ["run", "--directory", "<home>\\sharepoint-mcp", "server.py"]. The config is %APPDATA%\Claude\claude_desktop_config.json, or for the Microsoft Store version under %LOCALAPPDATA%\Packages\Claude_*\LocalCache\Roaming\Claude\. Merge into the existing file; don't overwrite other servers. Back up the file first. Escape backslashes in JSON.
   Use real absolute paths.
8. Ask me whether to enable list writes, and explain first: setting SHAREPOINT_ENABLE_WRITES=1 adds one tool, sp_list_item_update, which edits existing SharePoint list items (no create or delete) as my `az` account, and clients ask before each call. Default to no. Only if I say yes, add "SHAREPOINT_ENABLE_WRITES": "1" to the env of each config entry (for Claude Code, re-run the add command with `-e SHAREPOINT_ENABLE_WRITES=1` before `--`). Never set it without my yes.
9. Verify: `claude mcp list` shows sharepoint-mcp as connected. Tell me to fully quit Claude Desktop from the tray and reopen it if you edited its config.
10. Summarise what you installed, where the config entries are, and anything I still need to do.
````

### Setup by hand


1. Install OneDrive and sign in with the work account. In SharePoint, add
   each library you want the server to see with **Sync** or **Add shortcut to
   My files**.
2. `winget install astral-sh.uv` and `winget install Microsoft.AzureCLI`, then
   `az login`. The Azure CLI must be the MSI/winget install, which ships its
   own `python.exe`. A pip-installed `az` isn't supported.
3. Clone the repo and run `uv run test_server.py`.
4. Register the server with the full path to `uv.exe` (`where uv`). Claude
   Desktop's config is `%APPDATA%\Claude\claude_desktop_config.json`:

   ```json
   "sharepoint-mcp": {
     "command": "C:\\Users\\you\\.local\\bin\\uv.exe",
     "args": ["run", "--directory", "C:\\Users\\you\\sharepoint-mcp", "server.py"]
   }
   ```

   Windows GUI apps inherit the user `PATH`, so no `env.PATH` entry is needed.
