# Smart Inventory Management & Stochastic Demand Analysis Using Claude and Excel MCP

University project for the course **Generative AI for Engineers**.

We connect **Claude Desktop** to Excel files through the **Model Context Protocol (MCP)**, so Claude can
create and edit workbooks from plain-language instructions. On top of this we build an inventory
model where daily demand is a random variable (Normal distribution) instead of a constant.

## Project stages

| Stage | Content | Status |
|---|---|---|
| 1 | Claude Desktop + Excel MCP Server setup, connection test, repository | ✅ Done |
| 2 | 30-day random demand (Normal distribution), EOQ, Safety Stock, Reorder Point, tables and charts | ✅ Done |
| 3 | Presentation, Zoom recording, live demo (e.g. changing demand σ from 10 to 20) | ⏳ Planned |

## Repository layout

```
smart-inventory-eoq/
├── data/        Excel workbooks Claude works on (the only folder the MCP server may access)
│                  Inventory_Test.xlsx   Stage 1 connection test
│                  Inventory_Model.xlsx  Stage 2 inventory model
├── docs/        Project documentation and notes
│                  stage1-mcp-test.md, stage2-inventory-model.md
├── outputs/     Charts, exports and results
└── setup/       Setup helper script and an example Claude Desktop config
```

## Dependencies

- [Claude Desktop](https://claude.ai/download)
- [uv](https://docs.astral.sh/uv/) – runs the MCP server
- [Excel MCP Server](https://github.com/haris-musa/excel-mcp-server) (MIT license, used as a dependency;
  its source code is **not** copied into this repository). Tested with version 2.2.2.
- Microsoft Excel (optional, only to view the files; the MCP server does not need it)

## Setup (Windows)

**1. Install uv.** Open PowerShell and run:

```powershell
winget install --id=astral-sh.uv -e
```

Close PowerShell, open a **new** window, and check that both commands print a version:

```powershell
uv --version
uvx --version
```

**2. Clone this repository.**

```powershell
cd $HOME\Documents
git clone https://github.com/energysystemmeka-hue/smart-inventory-eoq.git
cd smart-inventory-eoq
```

**3. Download the MCP server once** (so the first start in Claude does not time out):

```powershell
uvx excel-mcp-server --help
```

**4. Fully quit Claude Desktop.** Right-click the Claude icon in the system tray (next to the clock) and
choose **Quit**. Closing the window is not enough: a running Claude Desktop overwrites the config file.

**5. Add the Excel MCP Server to Claude Desktop.** From the repository folder run:

```powershell
powershell -ExecutionPolicy Bypass -File .\setup\add-excel-mcp.ps1
```

The script finds `uvx` and Claude Desktop's config file, keeps your existing settings, and allows the
server to use only this repository's `data` folder. It prints the paths it used.

<details>
<summary>Manual setup instead of the script</summary>

In Claude Desktop open **Settings → Developer → Edit Config** and add an `mcpServers` block like the one in
[`setup/claude_desktop_config.example.json`](setup/claude_desktop_config.example.json). Replace the
`--allow-dir` value with the full path of your `data` folder, using double backslashes (`\\`).
If Claude cannot find `uvx`, use its full path (run `(Get-Command uvx).Source` in PowerShell).
Save the file while Claude is closed, then start Claude.

</details>

**6. Start Claude Desktop** and check **Settings → Developer**: the `excel` server should be listed as running.

**7. Test it.** In a new chat in the **Chat** tab, send:

```
Are you connected to the Excel MCP server? List the Excel files in the allowed folder with list_workbooks.
```

Claude should ask for permission to use the tool and then list the workbooks in `data`
(`Inventory_Test.xlsx` and `Inventory_Model.xlsx`).

## Troubleshooting

| Problem | Fix |
|---|---|
| `uv : The term 'uv' is not recognized` | Open a new PowerShell window after installing uv. Terminals that were open before the install do not see it. |
| Claude says it has no Excel tools | Claude was not fully restarted, or it overwrote the config. Quit Claude from the system tray, run step 5 again, then start Claude. |
| Claude cannot open a file | The server only allows files inside the `data` folder. Use paths relative to it, e.g. `Inventory_Test.xlsx`. |
| Changes do not appear / file is locked | Close the workbook in Excel before asking Claude to change it. While Excel has the file open, a write fails with "Access denied" (Turkish Windows: "Erişim engellendi"). |
| Safety stock, reorder point or total cost come back empty | The MCP server's calculation engine supports `NORM.S.INV` but not the older `NORMSINV`. Use `NORM.S.INV`. |

## Stage 1 test file

`data/Inventory_Test.xlsx` was created and then edited by Claude through MCP (see
[`docs/stage1-mcp-test.md`](docs/stage1-mcp-test.md)).

## Stage 2 inventory model

`data/Inventory_Model.xlsx` is a single-product inventory model with stochastic daily demand.
Full details, formulas and test results are in
[`docs/stage2-inventory-model.md`](docs/stage2-inventory-model.md).

### Worksheets

| Sheet | Content |
|---|---|
| **Inputs** | All parameters in one place: mean daily demand μ = 50, standard deviation σ = 10, 365 days/year, ordering cost S = 100 TL/order, holding cost H = 2 TL/unit/year, lead time L = 5 days, service level 95 %, and the z value `=NORM.S.INV(B8)`. Each parameter has a defined name (`Mu`, `Sigma`, `LeadTime`, …) that the formulas use. |
| **Demand** | 30 days of simulated daily demand, a reference line at μ, 30-day statistics (mean, standard deviation, min, max) and a line chart. |
| **Results** | Annual demand, EOQ, orders per year, cycle length, lead-time demand, safety stock, reorder point, annual cost breakdown, a check of safety stock based on the 30-day sample, and a column chart of lead-time demand, safety stock and reorder point. |

### Demand generation

Column B of the **Demand** sheet holds 30 fixed standard normal draws Z ~ N(0, 1). Daily demand is

```
Demand = MAX(0, ROUND(μ + σ · Z, 0))
```

Because the draws are fixed (no `RAND()`), the sample is reproducible, and changing μ or σ on the
**Inputs** sheet updates the whole model at once.

### Formulas

| Quantity | Formula |
|---|---|
| Annual demand | D = μ × 365 |
| Economic order quantity | Q* = √(2DS / H) |
| Safety stock | SS = z × σ × √L, with z = `NORM.S.INV(service level)` |
| Reorder point | ROP = μ × L + SS |
| Total annual cost | (D / Q*) × S + (Q* / 2) × H + SS × H |

With the default inputs: EOQ ≈ 1,350.9 units, SS ≈ 36.8 units, ROP ≈ 286.8 units, total annual
cost ≈ 2,775.41 TL.

The z value uses `NORM.S.INV`, not the older `NORMSINV`. Both give the same result in Excel, but the
Excel MCP Server's calculation engine only supports `NORM.S.INV`. With `NORMSINV`, MCP returned empty
values for safety stock, reorder point and total cost.

### MCP tests

All of these were done through Excel MCP tools:

- Read all three sheets, formulas and defined names.
- Changed `Inputs!B9` from `=NORMSINV(B8)` to `=NORM.S.INV(B8)`. After that, all results were
  calculated. The original file was backed up first.
- Sensitivity test: changed σ from 10 to 20, read the new results, then set σ back to 10 and
  confirmed that all results returned to their original values. EOQ stayed at 1,350.9 units, safety
  stock doubled from 36.8 to 73.6 units, the reorder point rose from 286.8 to 323.6 units, and total
  annual cost rose from 2,775.41 to 2,848.97 TL.

## Team

3 students (Industrial Engineering and Energy Systems Engineering).
