# sharepoint-mcp

Read-only-by-default MCP server for SharePoint on macOS: OneDrive-synced
SharePoint libraries read from disk, and SharePoint lists via Microsoft Graph.
The file tools need no Microsoft Graph consent and no Full Disk Access; the list
tools (`sp_lists`, `sp_list_items`) use the Azure CLI's delegated Graph token
instead, so they require a signed-in `az`. List writes exist too, but are
[opt-in](#writes-opt-in).

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

## Writes (opt-in)

> [!WARNING]
> Setting `SHAREPOINT_ENABLE_WRITES` to a truthy value (`1`, `true`, `yes`, `on`)
> registers three tools that change SharePoint list data, including delete.
> Without it, the write tools are not registered and the server is read-only.

```sh
SHAREPOINT_ENABLE_WRITES=1 uv run server.py
```

| Tool | What it does |
|---|---|
| `sp_list_item_create` | Create an item from a `fields` dict |
| `sp_list_item_update` | Change only the given fields on an item |
| `sp_list_item_delete` | Delete an item (moves it to the site recycle bin) |

`fields` is a dict of column values, e.g. `{"status": "In Progress"}`.
These register with `read_only_hint=False`; delete also sets
`destructive_hint=True` so clients prompt.

## Requirements

- macOS, with at least one SharePoint library synced through OneDrive (file tools)
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
  "command": "/opt/homebrew/bin/uv",
  "args": ["run", "--directory", "/absolute/path/to/sharepoint-mcp", "server.py"]
}
```

Use an absolute path to `uv` in GUI apps; they don't inherit your shell `PATH`.
