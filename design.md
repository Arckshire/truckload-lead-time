# Truckload Lead Time Optimization (TL LTO) — Design

This document captures every design decision agreed before the tool was built. It is the source of truth; if behavior differs in code, fix the code.

## 1. Goal

A Streamlit web app that ingests the p44 unified shipment export for truckload shipments and produces lane-level + carrier-level lead-time analytics. Sibling of the existing ocean LTO tool, with all statistical and ranking logic identical and only TL-specific data shape changes layered on top.

## 2. Input

Source file is the p44 unified shipment export, CSV or Excel. The tool accepts files at least up to about 100 MB without crashing. On load the parser drops every column outside the wanted set, downcasts dtypes implicitly via pandas, and caches the parse step with `@st.cache_data` so a re-run does not re-parse the same upload.

Mandatory columns: `Shipment ID`, `Current mode`, `Current carrier`, plus the Stop 1 family (`Stop 1`, `Stop 1 city`, `Stop 1 state`, `Stop 1 country`, `Stop 1 actual arrival time`, `Stop 1 actual departure time`). Stops 2 through 20 follow the same naming convention and are read opportunistically.

Filtering is hard at load time. Only rows where `Current mode == TRUCKLOAD` are kept. AIR, RAIL, OCEAN, UNKNOWN, and any other mode are dropped.

Tenant column is optional. If the export does not carry a Tenant column, the app inserts a placeholder tenant (`TL Tenant`) so the rest of the grouping logic works. Carrier name is read from `Current carrier`; carrier SCAC is left blank because the p44 unified export has no TL SCAC field.

## 3. Milestones

There are exactly four always-available milestones, regardless of stop count:

| Key | Name                  | Resolves to                          |
| --- | --------------------- | ------------------------------------ |
| A   | Origin Arrival        | Stop 1 actual arrival time           |
| B   | Origin Departure      | Stop 1 actual departure time         |
| C   | Destination Arrival   | Last-stop actual arrival time        |
| D   | Destination Departure | Last-stop actual departure time      |

"Last stop" is detected dynamically per shipment as the largest `N` in 1..20 where `Stop N` has a populated name. Blank, empty, and single-space cells are treated as missing. The user never selects intermediate stops as a journey endpoint.

## 4. Journey window picker

The user picks a start milestone and an end milestone from the four. Default is `B → C` (Origin Departure → Destination Arrival), pure line-haul transit. Chronological order is enforced: the end must come after the start in `A → B → C → D` order. Violations produce a clear UI error and computation is halted.

The journey window picker is decoupled from the Whole Journey toggle. Whole Journey is an independent checkbox that, when on, additionally computes every named segment for each shipment.

## 5. Lane definition

The lane string is built from city, state, and country components for each endpoint. City is title-cased, state is upper-cased, country is upper-cased ISO. Missing components are dropped cleanly (no stray commas, no `Unknown` placeholder). Format with all three present: `Romulus, MI, US`.

Whole Journey OFF: the lane is `<first stop endpoint> → <last stop endpoint>`. Middle stops are ignored for grouping. A 2-stop shipment and a 3-stop shipment with the same first and last cities are in the same lane.

Whole Journey ON: the lane is the full ordered chain `<Stop 1 endpoint> → <Stop 2 endpoint> → ... → <Last stop endpoint>`. Each unique chain is its own lane. The same two shipments above split into two distinct lanes.

A shipment is dropped from the analysis if either the first or the last endpoint is fully blank (no city, no state, no country). A partially-populated endpoint (e.g., state and country present but city missing) is rendered as the populated parts only and the shipment is retained.

## 6. Whole Journey ON additions

When Whole Journey is on the app computes, per shipment, every named segment:

- Origin Dwell = Stop 1 departure minus Stop 1 arrival.
- Stop N Dwell for each intermediate stop N = Stop N departure minus Stop N arrival.
- Destination Dwell = Last-stop departure minus Last-stop arrival.
- Origin → Stop 2, Stop 2 → Stop 3, ..., Stop N-1 → Destination = transit between adjacent stops.
- Total Lead Time = end timestamp minus start timestamp (the picked journey window).

These segments use uniform "Dwell" naming. Internal column names are short semantic identifiers (`DWELL_1`, `SEG_1_2`, etc.); the human labels are resolved per-row from the row's last-stop index and surfaced in the export and the UI.

## 7. Data quality rules

If a segment computes to a negative duration on a particular shipment, that segment is excluded for that shipment but the shipment is still counted in every other segment. This is "per-metric exclusion" rather than "whole-shipment exclusion." Maximizes retained data without poisoning the stats with corrupted timestamps.

Eligibility for the default-journey metric requires only the chosen start and end milestone timestamps to be present. Eligibility for the Whole Journey metrics requires the corresponding stop arrival and departure timestamps for that segment to be present; absent timestamps simply produce a NaN for that segment, with no cascade to other segments.

## 8. Calculations (zero changes from ocean)

Per group statistics: count, median, P25, P75, configurable PXX (default 80), min, max, and total. Hours are rounded to two decimal places; days are hours divided by twenty-four rounded to the nearest integer.

Carrier ranking inside a lane is by lowest median absolute deviation from the lane median, then lowest PXX absolute deviation from the lane median, then higher shipment volume, then alphabetical carrier name. This is the predictability-first ranking inherited verbatim from the ocean tool. The wording "least deviated carrier" is the right mental model; the fastest carrier is not necessarily ranked first.

Insights output per metric is a two-tier structure: a lane summary (one row per lane with shipment count, lane median, lane PXX, carrier count), and carrier recommendations (one row per carrier-in-lane with shipments, share percent, carrier median, carrier PXX, median absolute deviation, PXX absolute deviation, rank in lane, and eligibility flags). The UI surfaces top 5 carriers per lane plus a deviation bar chart and a lead-time-by-carrier box plot.

## 9. Sidebar controls

The sidebar matches the ocean tool. Controls are: file upload (accepts CSV or Excel), journey start milestone, journey end milestone, Whole Journey checkbox, Top-N lanes filter, percentile checkbox, percentile value, percentile volume threshold (checkbox + input, off by default), recommendation volume threshold (checkbox + input, off by default).

Volume thresholds are percentage-based against lane shipment counts. By default both thresholds are off so a one-shipment lane is retained in all output.

## 10. Main panel

In-page preview tables are capped at top 25 rows. Counts (lane and carrier), the shipment-level lead-time table, the carrier lane lead report, and the insights tables all preview at most 25 rows. Full data is always available via the download buttons.

A "Generate Insights" button reveals the insights section. The lane selector inside the insights section is keyed off the lane summary. Inside the insights section the user can also flip between the metrics (journey total in default mode, or every named segment plus total in Whole Journey mode).

Progress is shown via Streamlit's `st.status` and `st.progress` primitives during the file read and the shipment compute phases. Errors are surfaced via `st.error` for blocking failures, `st.warning` for empty result states, and `st.info` for informational notices like "filtered out N non-TL rows."

## 11. Exports

Excel export uses one workbook with multiple sheets: Raw Data, Carrier Lane Lead, and Key (glossary). For the counts download the sheets are Lane Counts, Carrier Counts, and Key. For the insights download the sheets are Lane Summary, Carrier Recommendations, Selected Lane Shipments, and Key. Headers are bolded and column widths are sized for readability.

CSV export packages the same content as a single ZIP file containing one CSV per sheet plus `key.csv`. Both formats are offered side by side everywhere a download button appears.

File names follow the ocean convention. Counts download is `lane_and_carrier_counts.xlsx` / `.zip`. Final report is `tl_carrier_lane_lead_<start>_to_<end>.xlsx` / `.zip`. Insights download is `tl_insights_<metric-label>.xlsx` / `.zip` where the metric label is sanitized for filename safety.

## 12. Robustness

The CSV parser tries UTF-8, UTF-8-with-BOM, CP1252, and Latin-1 in that order. The Excel parser uses pandas' default openpyxl engine. Both parsers read the header first to compute the wanted column subset, then re-read with `usecols` to load only what we need. The parse is wrapped in `@st.cache_data` keyed on the raw bytes and the file name, so toggling sidebar controls does not retrigger the parse.

The compute path uses runtime calculation on every interaction, the same pattern ocean uses. No precomputed in-memory cache of segment durations. The only caching is around the file parsing step, which is invisible to the user and does not touch any business logic.

## 13. Versioned scope

Version 1 does not compute planned-vs-actual variance. Planned and appointment timestamps are passed through in the raw export only. Version 1 does not provide route-level analysis as a separate concept; in TL terms, "route insight" collapses into "whole journey lane definition" — each unique stop chain is its own lane when Whole Journey is on.

Out of scope for v1: Snowflake direct connectivity (the platform CSV/Excel export is the source), multi-tenant comparison views, side-by-side default vs Whole Journey results, scheduled refreshes.
