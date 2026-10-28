# Live Demonstration Guide

**Smart Inventory Management & EOQ Analysis Using Claude and Excel MCP**

This guide is the script for a 5–7 minute live demonstration. In the demo, **Claude Code** runs in
**PowerShell** and reads and changes `data/Inventory_Model.xlsx` through the Excel MCP Server
(`excel-mcp-server`). Every number in this guide was measured through Excel MCP during Stage 2
testing (see [`stage2-inventory-model.md`](stage2-inventory-model.md)).

---

## 1. Before the demonstration (checklist)

Do this 15–30 minutes before presenting. Run the commands in PowerShell in the repository folder.

| # | Check | How |
|---|---|---|
| 1 | You are in the repository folder | `cd` to the folder that contains `README.md` and `data`. The Excel server is configured for this project, so Claude Code must be started here. |
| 2 | The Excel MCP server is connected | Run `claude mcp list`. The `excel` server must be shown as connected (see below). |
| 3 | The workbook has no unexpected changes | Run `git status`. `data/Inventory_Model.xlsx` should not be listed as modified. If it is, see section 8 before doing anything else. |
| 4 | Excel is **closed** | No Excel window is open, and there is no hidden lock file `data\~$Inventory_Model.xlsx` (check with `Get-ChildItem -Force data`). |
| 5 | Start a fresh session | Run `claude` to start a new Claude Code session, so earlier context does not influence the answers. |
| 6 | Dry run | Run prompts 1 and 2 (section 3) once. Check that σ = 10 and safety stock = 36.78. Then exit (`/exit`) and start a new session for the real demo. |
| 7 | Screen layout | PowerShell with Claude Code on the main screen, with a large font. File Explorer open at the `data` folder, ready to open the workbook. |

Expected line in the output of `claude mcp list` (the path is your own `data` folder):

```
Checking MCP server health…

excel: uvx excel-mcp-server stdio --allow-dir <repository>\data - ✔ Connected
```

Other servers may also be listed, depending on your account and plugins. Only the `excel` line
matters for the demo. If it shows "Failed to connect" or is missing, see section 8.

**Permission prompts.** Claude Code asks for permission before it runs an MCP tool that has not been
allowed yet. In the demo, approve each request (*Yes*). These prompts are useful: they show the
audience the exact tool and arguments before anything happens.

---

## 2. Demonstration plan (about 6 minutes)

| Time | Step | Who / what | Prompt |
|---|---|---|---|
| 0:00 – 0:45 | Introduction: the problem and the tools | Speaker | – |
| 0:45 – 1:30 | Show the MCP connection | PowerShell + Claude Code | `claude mcp list`, `/mcp`, prompt 1 |
| 1:30 – 2:45 | Read the model and the current results (σ = 10) | Claude Code | Prompt 2 |
| 2:45 – 4:00 | Sensitivity analysis: change σ from 10 to 20 (Excel closed) | Claude Code | Prompt 3 |
| 4:00 – 4:45 | Open the workbook in Excel to show the result, then close it **without saving** | Speaker | – |
| 4:45 – 5:30 | Restore σ = 10 and verify the original results (Excel closed) | Claude Code | Prompt 4 |
| 5:30 – 6:30 | Interpretation and engineering value | Speaker | – |

Suggested introduction (0:00–0:45):

> "Inventory planning usually assumes constant demand. Real demand varies. We built a model in
> Excel where daily demand is Normally distributed, and we connected Claude Code to Excel through
> the Model Context Protocol. Claude does not guess the numbers: it reads and writes the workbook
> with Excel tools, and the workbook's formulas do the calculation."

---

## 3. Prompts for the demonstration

### Showing the connection (0:45)

1. Before starting Claude Code, run in PowerShell:

   ```powershell
   claude mcp list
   ```

2. Start Claude Code with `claude`, then type `/mcp`. Select the `excel` server to show its status
   and its list of tools (`list_workbooks`, `describe_workbook`, `read_range`, `write_range`, …).
   Press `Esc` to return.

### Rules used in every prompt

Claude Code can also run shell commands and scripts. To show that the **Excel MCP tools** are doing
the work, every prompt says explicitly that only these tools may be used. The tools take paths
relative to the `data` folder, so the file name is just `Inventory_Model.xlsx`.

Tools used in the demo:

| Tool | Purpose |
|---|---|
| `list_workbooks` | Lists the Excel files in the allowed folder |
| `describe_workbook` | Sheets, used ranges, charts, defined names |
| `read_range` | Reads cells; `mode` = `values` (results) or `formulas` (formula text) |
| `write_range` | Writes values or formulas into cells |

### Prompt 1 – Connection

```
Use only the Excel MCP tools, no shell commands or scripts.
Call list_workbooks to list the Excel files in the allowed folder. Then call describe_workbook
on Inventory_Model.xlsx and list its sheets, its charts and its defined names.
Do not change anything.
```

Expected:

- Files: `Inventory_Model.xlsx`, `Inventory_Test.xlsx`, and `Inventory_Model_backup.xlsx` if the
  backup still exists in the folder.
- Sheets: **Inputs**, **Demand**, **Results**. Demand and Results each have one chart.
- 13 defined names, for example `Mu`, `Sigma`, `LeadTime`, `z_value`, `SafetyStock`,
  `ReorderPoint`, `TotalCost`.

### Prompt 2 – Read the model (σ = 10)

```
Use only the Excel MCP tools, no shell commands or scripts. Do not change anything.
Call read_range on Inventory_Model.xlsx:
- sheet Inputs, range A1:C9, mode formulas
- sheet Results, range A1:B12, mode values
Show the input parameters, the formula in Inputs!B9, and the values of EOQ, Safety Stock,
Reorder Point and Total Annual Cost.
```

Expected (see section 4, column σ = 10). The formula in `Inputs!B9` is `=NORM.S.INV(B8)`.

### Prompt 3 – Sensitivity analysis (σ = 20)

Check first that Excel is closed.

```
Use only the Excel MCP tools, no shell commands or scripts.
The workbook is closed in Excel. In Inventory_Model.xlsx, call write_range to change ONLY
Inputs!B3 (Sigma) from 10 to 20. Do not change any other cell.
Then call read_range on Results!A1:B12 and on Demand!F2:G5 (mode values).
Show a table that compares EOQ, Safety Stock, Reorder Point and Total Annual Cost for
Sigma = 10 and Sigma = 20, and explain the changes briefly.
```

Expected: see section 4. EOQ stays the same, safety stock doubles, and the reorder point and total
cost rise.

### Between prompts 3 and 4 – Show the workbook in Excel

1. Open `data\Inventory_Model.xlsx` in Excel. This is only for display; MCP has already saved the
   change.
2. Show **Inputs!B3 = 20**, the new safety stock and reorder point on **Results**, and the wider
   demand line in the chart on **Demand**.
3. Close Excel and choose **Don't Save**.
4. Check that the lock file has gone before the next write: `Get-ChildItem -Force data` should not
   list `~$Inventory_Model.xlsx`.

### Prompt 4 – Restore and verify

```
Use only the Excel MCP tools, no shell commands or scripts.
The workbook is closed in Excel. Call write_range to set Inputs!B3 in Inventory_Model.xlsx back
to 10, changing no other cell. Then call read_range on Inputs!B3 and on Results!A1:B12
(mode values) and confirm that EOQ, Safety Stock, Reorder Point and Total Annual Cost are back
to their original values.
```

Expected: σ = 10, safety stock 36.78, reorder point 286.78, total annual cost 2,775.41 TL.

---

## 4. Expected numerical results

All values were read through Excel MCP. The fixed inputs are μ = 50 units/day,
365 days/year, S = 100 TL/order, H = 2 TL/unit/year, L = 5 days and service level = 0.95, which
gives z = 1.6449.

| Metric | σ = 10 | σ = 20 | Change |
|---|---|---|---|
| Annual demand D | 18,250 | 18,250 | none |
| **EOQ Q\*** | **1,350.93** | **1,350.93** | **none** |
| Orders per year | 13.51 | 13.51 | none |
| Cycle length (days) | 27.02 | 27.02 | none |
| Lead-time demand | 250 | 250 | none |
| **Safety stock** | **36.78** | **73.56** | **×2** |
| **Reorder point** | **286.78** | **323.56** | **+36.78 (+12.8 %)** |
| Annual ordering cost (TL) | 1,350.93 | 1,350.93 | none |
| Annual cycle holding cost (TL) | 1,350.93 | 1,350.93 | none |
| Annual safety stock cost (TL) | 73.56 | 147.12 | ×2 |
| **Total annual cost (TL)** | **2,775.41** | **2,848.97** | **+73.56 (+2.65 %)** |
| 30-day sample std. deviation | 10.03 | 20.05 | ×2 |
| 30-day sample min / max | 27 / 73 | 3 / 95 | wider range |

Claude may round the values differently, for example 1,350.9 or 1,351. That is expected.

---

## 5. Explanation for the audience

**EOQ – how much to order.**
Q\* = √(2DS / H) = √(2 × 18,250 × 100 / 2) ≈ 1,351 units. EOQ is the order size at which annual
ordering cost and annual holding cost are equal, here 1,350.93 TL each. It depends on average
demand, not on the variation of demand. This is why it does not change when σ goes from 10 to 20.

**Safety stock – protection against variation.**
SS = z × σ × √L = 1.6449 × 10 × √5 ≈ 36.8 units. This is the extra stock that keeps the
probability of no stockout during the 5-day lead time at 95 %. z comes from the service level with
`NORM.S.INV(0.95)`. Safety stock is proportional to σ, so doubling σ doubles it (73.56 units).

**Reorder point – when to order.**
ROP = μ × L + SS = 250 + 36.78 ≈ 286.8 units. When stock falls to this level, a new order is
placed. 250 units are expected to be used during the lead time, and the safety stock covers
higher-than-expected demand. With σ = 20 the ROP rises to 323.56, so orders are placed earlier.

**Total annual cost.**
Ordering cost + cycle holding cost + safety stock holding cost. Only the last part depends on σ, so
doubling the demand variation raises total cost by only 73.56 TL per year (+2.65 %).

Key message: **higher uncertainty does not change how much we order. It changes when we order and
how much buffer stock we keep.**

---

## 6. Excel must be closed while MCP writes

Excel locks a workbook while it is open. If Excel has `Inventory_Model.xlsx` open, an MCP write
fails with "Access denied" (Turkish Windows: "Erişim engellendi"). Even if a write succeeded, Excel
would still hold the old version and could overwrite the change when it saves.

| Phase | Excel | Claude Code / MCP |
|---|---|---|
| Read (prompts 1, 2) | Closed (recommended) | Reads |
| **Write (prompts 3, 4)** | **Must be closed** | Writes and saves the file |
| Show the result | **Open afterward** to display the results, then close with **Don't Save** | Idle |

Rules:

1. **Excel must be closed while MCP writes to the workbook.** Open it only after the write has
   finished, to display the results.
2. **Close Excel with "Don't Save".** MCP has already saved the change. Saving from Excel is not
   needed, and it writes the Windows/Office user name into the file's properties.
3. Before the next write, check that the hidden lock file `data\~$Inventory_Model.xlsx` has gone
   (`Get-ChildItem -Force data`).

---

## 7. Showing that Excel MCP is really used

Show these points to the audience:

1. **`claude mcp list` and `/mcp`.** They show the `excel` server as connected, its command
   (`uvx excel-mcp-server stdio --allow-dir …\data`) and its tools. The server can access only the
   project's `data` folder.
2. **Tool calls in the session.** Claude Code shows each call with the server and tool name, for
   example `excel - read_range (MCP)`, followed by its arguments (file, sheet, range, mode) and its
   result. Press `Ctrl+O` to expand the full tool output if it is collapsed.
3. **Permission prompts.** Before a tool runs for the first time, Claude Code asks for permission
   and shows the tool name and arguments.
4. **No shell commands.** Every prompt forbids shell commands and scripts, and the audience can see
   that only `excel` tools are called.
5. **The formula, not only the number.** Prompt 2 reads `Inputs!B9` in formulas mode and returns
   `=NORM.S.INV(B8)`. This text exists only in the workbook.
6. **The change appears in Excel.** After prompt 3, open the file in Excel: `Inputs!B3` is 20 and
   the results on the Results sheet have changed, although nobody typed in Excel.
7. **The file on disk changed.** Optional: in a second PowerShell window, run `git status` after
   prompt 3. `data/Inventory_Model.xlsx` is listed as modified. After prompt 4 the cell values are
   the same as before, but Git may still list the file as modified, because MCP saved the file
   again. This is expected.

---

## 8. Troubleshooting

| Problem | Likely cause | Solution |
|---|---|---|
| "Access denied" / "Erişim engellendi" on a write | The workbook is open in Excel. | Close Excel (Don't Save), check that `~$Inventory_Model.xlsx` is gone, and send the prompt again. A failed write changes nothing. |
| `claude mcp list` does not show `excel`, or shows "Failed to connect" | Claude Code was started in another folder, or `uv` / `uvx` is not available. | Run the command from the repository folder. Check `uvx --version` and `uvx excel-mcp-server --help`. Run `claude mcp get excel` to see the configured command and `--allow-dir` path. |
| `/mcp` shows the server, but Claude says it has no Excel tools | The session started before the server connected. | Exit (`/exit`) and start `claude` again from the repository folder. |
| Claude uses shell commands or a script instead of the Excel tools | The prompt did not restrict the tools. | Deny the permission request, and repeat the prompt starting with "Use only the Excel MCP tools, no shell commands or scripts." |
| Claude answers without calling a tool | The model answered from earlier context. | Start a new session, or add "Call read_range now to read the workbook" to the prompt. |
| Safety stock, ROP or total cost are empty | `Inputs!B9` contains the old `NORMSINV`, which the MCP calculation engine does not support. This happens if the old backup file is used. | Check `Inputs!B9` with `read_range` in formulas mode. It must be `=NORM.S.INV(B8)`. Do **not** copy `Inventory_Model_backup.xlsx` over the model: the backup still uses `NORMSINV`. |
| Results are not the expected values | σ or another input was left changed, for example after a rehearsal. | Read `Inputs!A1:C9` with `read_range`. Set the changed input back with `write_range` (for σ, use prompt 4) and read the results again. |
| Excel asks to save when closing | Excel recalculates the workbook when it opens. | Choose **Don't Save**. |
| `git status` lists `Inventory_Model.xlsx` as modified before the demo | A rehearsal, an MCP write or a save from Excel changed the file. | See "Resetting the workbook" below. |
| Numbers are rounded differently | Claude formats the output. | No action needed; the stored values do not change. |

### Resetting the workbook

Prefer to fix the workbook with MCP: read `Inputs!A1:C9` and set any changed input back with
`write_range`, then compare the results with section 4.

Avoid `git checkout -- data/Inventory_Model.xlsx` (or `git restore`) as a quick reset. It
**permanently discards all uncommitted changes** to the workbook, including changes you may want to
keep, and they cannot be recovered from Git afterward. Use it only after you have:

1. checked with `git status` and `git diff --stat` what is modified;
2. confirmed with `read_range` that the changes are not needed; and
3. made a copy of the current file outside the `data` folder if there is any doubt.

### Fallback

If the live connection fails, present section 4 of this guide together with the sensitivity
analysis in [`stage2-inventory-model.md`](stage2-inventory-model.md). All values there were
measured through Excel MCP.

---

## 9. Conclusion – engineering value

- **Uncertainty is modeled, not ignored.** Daily demand follows a Normal distribution, and safety
  stock and reorder point are derived from a chosen service level instead of a rule of thumb.
- **Decisions are separated clearly.** The demonstration shows that demand variation affects *when*
  to order (ROP) and how much buffer to keep (SS), but not *how much* to order (EOQ).
- **What-if analysis takes seconds.** One request in plain language changes σ, and the whole model
  updates. This makes sensitivity analysis available to people who do not edit formulas themselves.
- **Results stay traceable.** Claude works through defined tools on a real workbook. Every input,
  formula and result can be checked in Excel, and every change can be checked with Git.
- **The integration was tested.** During testing we found that the MCP calculation engine does not
  support the old `NORMSINV` function, changed the model to `NORM.S.INV`, and verified the result.
  AI-assisted engineering tools must be verified, not assumed to be correct.
