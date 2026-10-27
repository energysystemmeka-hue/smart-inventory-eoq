# Stage 1 – MCP connection test

Goal: show that Claude Desktop can create and edit an Excel file through the Excel MCP Server.

Environment: Windows 11, Claude Desktop, uv 0.13.6, Excel MCP Server 2.2.2 (46 tools).
The server was started with `--allow-dir <repo>/data`, so it can only access the `data` folder.

## 1. Connection check

Prompt (Chat tab):

> Are you connected to the Excel MCP server? List the Excel files in the allowed folder with list_workbooks.

Result: Claude called `list_workbooks`; the folder was empty, so the list was empty.

## 2. Create the workbook

Prompt:

> Using the Excel MCP tools, create a new file named Inventory_Test.xlsx with a sheet named "Inventory".
> Headers in A1:D1: Product, Daily Demand, Ordering Cost, Holding Cost. Add 3 sample products:
> Product A, 50, 100, 2 / Product B, 30, 80, 1.5 / Product C, 20, 120, 3. Make the header row bold.
> Then read A1:D4 with read_range and show it.

Result (checked independently by reading the file from disk with `openpyxl`):

| Product | Daily Demand | Ordering Cost | Holding Cost |
|---|---|---|---|
| Product A | 50 | 100 | 2 |
| Product B | 30 | 80 | 1.5 |
| Product C | 20 | 120 | 3 |

Sheet `Inventory`, used range A1:D4, header row bold.

## 3. Edit a cell

Prompt:

> In Inventory_Test.xlsx, sheet Inventory, change B3 (Product B's Daily Demand) from 30 to 45.
> Do not touch any other cell. Then read A1:D4 again.

Result (checked from disk): B3 = 45; all other cells and the formatting unchanged.

## Lessons learned

- After installing uv, terminals and apps that were already open do not see it; open a new window.
- The Store (MSIX) build of Claude Desktop keeps its config under
  `%LOCALAPPDATA%\Claude-Data\claude_<id>\claude_desktop_config.json`.
- If Claude Desktop is running while the config is edited, it can overwrite the file and drop the
  `mcpServers` block. Edit the config only while Claude is fully quit (see `setup/add-excel-mcp.ps1`).
