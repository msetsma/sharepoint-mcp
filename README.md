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

`sp_lists` and `sp_list_items` borrow the `az` CLI's delegated Graph token, so
run `az login` once before using them (no app registration needed).

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

- macOS or Windows, with at least one SharePoint library synced through OneDrive (file tools)
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

## Windows

The server runs on Windows with the same tools. Three things differ from
macOS, and the code handles each one:

| | macOS | Windows |
|---|---|---|
| Where synced libraries live | `~/Library/CloudStorage/OneDrive-SharedLibraries-<Org>` | `%USERPROFILE%\<Org>\<Site> - <Library>`, read from OneDrive's registry entries under `HKCU\Software\SyncEngines\Providers\OneDrive` (personal OneDrive is skipped) |
| How a cloud-only file is spotted | `st_blocks == 0` | `FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS` or `FILE_ATTRIBUTE_OFFLINE` set |
| How `az` is launched | the `az` script | the CLI's bundled `python.exe -IBm azure.cli`, bypassing `az.cmd` so `cmd.exe` never re-parses `&`, `%` or `"` in URLs and JSON bodies |

Setup:

1. Install OneDrive and sign in with the work account. In SharePoint, click
   **Sync** on each library you want the server to see.
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
