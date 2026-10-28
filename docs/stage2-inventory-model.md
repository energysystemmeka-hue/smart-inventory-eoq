# Stage 2 – Stochastic inventory model (EOQ, Safety Stock, Reorder Point)

Workbook: `data/Inventory_Model.xlsx` (sheets **Inputs**, **Demand**, **Results**).

## 1. Objective

Build an inventory model in Excel for one product whose daily demand is a random variable instead
of a constant, and work on it through Claude and the Excel MCP Server. The model answers two
questions:

- **How much to order?** The economic order quantity (EOQ).
- **When to order?** The reorder point (ROP), which includes a safety stock that protects against
  demand variation during the lead time.

It also shows how the results change when demand becomes more variable (σ = 10 versus σ = 20).

## 2. Assumptions and parameters

Assumptions:

- Single product, continuous review: an order of size Q* is placed when stock falls to the ROP.
- Daily demand is independent and Normally distributed, D_day ~ N(μ, σ²).
- Lead time is constant.
- Ordering cost per order and holding cost per unit per year are constant.
- No quantity discounts. Shortage cost is not modeled; protection against shortages is set by the
  service level.
- Service level means the probability of no stockout during the lead time (cycle service level).

Parameters (sheet **Inputs**, each with a defined name):

| Cell | Parameter | Name | Value | Unit |
|---|---|---|---|---|
| B2 | Mean daily demand | `Mu` | 50 | units/day |
| B3 | Standard deviation of daily demand | `Sigma` | 10 | units/day |
| B4 | Days per year | `DaysPerYear` | 365 | days |
| B5 | Ordering cost (S) | `OrderCost` | 100 | TL/order |
| B6 | Holding cost (H) | `HoldCost` | 2 | TL/unit/year |
| B7 | Lead time (L) | `LeadTime` | 5 | days |
| B8 | Service level | `ServiceLevel` | 0.95 | – |
| B9 | z value | `z_value` | `=NORM.S.INV(B8)` → 1.644854 | – |

## 3. Formulas

Sheet **Results**, column B:

| Cell | Metric | Formula | Excel formula |
|---|---|---|---|
| B2 | Annual demand | D = μ × 365 | `=Mu*DaysPerYear` |
| B3 | EOQ | Q* = √(2DS / H) | `=SQRT(2*B2*OrderCost/HoldCost)` |
| B4 | Orders per year | N = D / Q* | `=B2/B3` |
| B5 | Cycle length | T = Q* / μ (days) | `=B3/Mu` |
| B6 | Lead-time demand | μ × L | `=Mu*LeadTime` |
| B7 | Safety stock | SS = z × σ × √L | `=z_value*Sigma*SQRT(LeadTime)` |
| B8 | Reorder point | ROP = μ × L + SS | `=B6+B7` |
| B9 | Annual ordering cost | (D / Q*) × S | `=B4*OrderCost` |
| B10 | Annual cycle holding cost | (Q* / 2) × H | `=B3/2*HoldCost` |
| B11 | Annual safety stock cost | SS × H | `=B7*HoldCost` |
| B12 | Total annual cost | B9 + B10 + B11 | `=B9+B10+B11` |
| B14 | Check: SS from the 30-day sample | z × sample sd × √L | `=z_value*Demand!G3*SQRT(LeadTime)` |

σ is a daily value and L is in days, so σ√L is the standard deviation of total demand during the
lead time.

### Why NORM.S.INV

`NORM.S.INV(p)` returns the z value with P(Z ≤ z) = p for the standard normal distribution. The
workbook first used the older name `NORMSINV`. Excel calculates both the same way, but the Excel MCP
Server (2.2.2) calculates formulas with its own engine, and that engine supports only `NORM.S.INV`.
With `NORMSINV`, MCP returned empty values for B9 and for every cell that depends on it (safety
stock, reorder point, safety stock cost, total cost and the sample check). After the change, all of
them are calculated. The z value stays linked to the service level, so changing B8 updates z and
everything after it.

## 4. Demand generation

Sheet **Demand**, rows 2–31 (days 1–30):

| Column | Content |
|---|---|
| A | Day 1–30 |
| B | Fixed standard normal draw Z ~ N(0, 1), stored as a number |
| C | Daily demand `=MAX(0,ROUND(Mu+Sigma*B,0))` |
| D | Mean line `=Mu` |
| F–G | 30-day statistics: `AVERAGE`, `STDEV`, `MIN`, `MAX` of column C |

This is the standard transformation X = μ + σZ, which turns a standard normal draw into a draw from
N(μ, σ²). The result is rounded to whole units, and `MAX(0, …)` prevents negative demand.

The Z values are stored as fixed numbers instead of `RAND()` formulas. So:

- the sample is reproducible, and the workbook shows the same demand every time it is opened;
- a change of μ or σ rescales the same 30 draws, so a comparison between σ values is not affected
  by a new random sample.

A line chart (`DemandChart`) shows daily demand (column C) against the mean (column D). A column
chart on the Results sheet (`ROPChart`) shows lead-time demand, safety stock and reorder point
(B6:B8).

## 5. Sigma = 10 versus Sigma = 20

Values read through Excel MCP after changing only `Inputs!B3`. All other inputs were unchanged.

| Metric | σ = 10 | σ = 20 | Change |
|---|---|---|---|
| Annual demand D | 18,250 | 18,250 | none |
| EOQ Q* | 1,350.93 | 1,350.93 | none |
| Orders per year | 13.51 | 13.51 | none |
| Cycle length (days) | 27.02 | 27.02 | none |
| Lead-time demand | 250 | 250 | none |
| Safety stock | 36.78 | 73.56 | ×2 |
| Reorder point | 286.78 | 323.56 | +36.78 (+12.8 %) |
| Annual ordering cost (TL) | 1,350.93 | 1,350.93 | none |
| Annual cycle holding cost (TL) | 1,350.93 | 1,350.93 | none |
| Annual safety stock cost (TL) | 73.56 | 147.12 | ×2 |
| Total annual cost (TL) | 2,775.41 | 2,848.97 | +73.56 (+2.65 %) |
| Sample mean (30 days) | 49.93 | 49.90 | – |
| Sample standard deviation | 10.03 | 20.05 | ×2 |
| Sample min / max | 27 / 73 | 3 / 95 | wider range |
| Check: SS from sample | 36.88 | 73.73 | ×2 |

## 6. Interpretation

- **EOQ does not change.** Q* = √(2DS / H) depends on the mean demand D, the ordering cost S and the
  holding cost H, not on σ. The order quantity balances ordering and holding costs, and these depend
  on average demand only. At Q*, annual ordering cost equals annual cycle holding cost
  (1,350.93 TL each), as the EOQ model predicts.
- **Safety stock doubles.** SS = z σ √L grows linearly with σ. With z and L fixed, doubling σ
  doubles the buffer that is needed to keep the 95 % service level during the 5-day lead time.
- **The reorder point rises by the same amount as safety stock.** Expected lead-time demand stays
  250 units, so the ROP increases from 286.78 to 323.56 units. Orders must be placed earlier.
- **Total cost rises only slightly.** The extra cost is the holding cost of the extra safety stock,
  36.78 × 2 = 73.56 TL per year (+2.65 %). Ordering and cycle holding costs do not change.
- **The sample matches the theory.** The 30-day sample standard deviation is close to the input σ
  (10.03 and 20.05), and safety stock based on the sample is within 0.3 % of the theoretical value.

In short, higher demand uncertainty does not change how much to order. It changes when to order
and how much buffer stock to keep.

## 7. MCP testing and verification

All workbook operations below were done through Excel MCP Server tools (`describe_workbook`,
`describe_sheet`, `read_range`, `trace_formula`, `write_range`).

1. **Inspection.** Read the three sheets in value and formula mode, the defined names and the
   charts. `read_range` reported B9 and its dependent cells as not calculated, with the reason
   `NORMSINV`. `trace_formula` gave the same error and listed the dependent cells
   (Results!B7:B8, B11:B12, B14).
2. **Cause.** The workbook file contains no stored formula results (it was not saved by Excel), so
   the MCP server calculates every formula with its own engine. That engine implements
   `NORM.S.INV`, `NORM.S.DIST`, `NORM.INV` and `NORM.DIST`, but not `NORMSINV`.
3. **Fix.** A byte-identical backup `data/Inventory_Model_backup.xlsx` was made first. Then only
   `Inputs!B9` was changed to `=NORM.S.INV(B8)` with `write_range`. The server stores the formula
   as `_xlfn.NORM.S.INV(B8)`, which is how Excel stores this function. A comparison with the backup
   showed that B9 was the only changed cell.
4. **Verification.** Through MCP: z = 1.644854, SS = 36.78, ROP = 286.78, total annual cost =
   2,775.41 TL.
5. **Sensitivity test.** Set σ to 20, read the results (section 5), set σ back to 10, and confirmed
   that all results returned to the original values.

Issue found during testing: the first attempt to write σ = 20 failed with "Access denied" because
the workbook was open in Microsoft Excel. Nothing was changed by the failed attempt. After Excel was
closed, the write worked.

## 8. Known limitations

- **One fixed sample.** The 30 Z values are a single fixed sample. The model shows one possible
  30-day path, not many simulation runs.
- **Rounding and truncation.** Rounding to whole units and `MAX(0, …)` change the distribution
  slightly. With σ = 20 the smallest simulated demand is 3, so the lower limit at 0 is not reached
  in this sample, but it could be for larger σ.
- **Planning uses parameters, not the sample.** EOQ, safety stock and ROP use the input μ and σ.
  The sample is used only for the check in Results!B14.
- **Constant lead time.** Lead-time variability is not modeled. If lead time varied, the safety
  stock formula would need an extra term.
- **No shortage cost.** The service level is set as a policy; the cost of a stockout is not in the
  total cost.
- **No stored results in the file.** The workbook has no cached formula values. Excel recalculates
  on opening (`fullCalcOnLoad`), and the MCP server calculates the values itself. Software that
  reads only stored values sees empty cells until the file has been saved by Excel.
- **MCP function support.** The MCP calculation engine does not support every Excel function name.
  New formulas should be checked through MCP after they are written.
- **File locking.** The workbook must be closed in Excel while Claude changes it through MCP.
- **Not checked in this stage:** the appearance of the charts in Excel after the `NORM.S.INV`
  change.
