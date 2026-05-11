# Truckload Lead Time Optimization Tool

A Streamlit web app for analyzing truckload (TL) shipment lead times and ranking carriers by lane predictability. Sibling tool to the existing ocean LTO tool, with identical statistical and ranking logic adapted for the dynamic-stop-count nature of TL data.

## What it does

Upload a p44 unified shipment CSV or Excel export. The app filters to truckload-only rows, builds lane strings from stop city/state/country, computes lead-time statistics per lane and per carrier, and ranks carriers within each lane by predictability (median absolute deviation from the lane median). Output is browsable in the UI and exportable as a multi-sheet Excel file or a ZIP of CSVs.

## Key features

End-to-end shipment lead-time calculation between any two of four always-available milestones: Origin Arrival, Origin Departure, Destination Arrival, Destination Departure. The "destination" milestone resolves to the dynamically detected last populated stop on each shipment, so the same picker works for 2-stop and 20-stop shipments.

Optional Whole Journey mode that decomposes each shipment into named segments — Origin Dwell, every intermediate stop dwell, every adjacent transit, the Destination Dwell, and the total — and treats each unique stop chain as its own lane.

Carrier ranking by lowest median absolute deviation from the lane median (with PXX deviation, volume, and alphabetical tiebreakers). Configurable PXX percentile (default 80) and optional percentage-based volume thresholds for both percentile inclusion and recommendation eligibility.

Insights tab that surfaces top 5 carriers per lane with a deviation bar chart and a lead-time-by-carrier box plot referenced against the lane median.

Multi-sheet Excel export and ZIP-of-CSVs export at every download point, with a Key/glossary sheet describing every column.

## Requirements

Python 3.9 or later. Dependencies are pinned in `requirements.txt`:

```
streamlit>=1.36.0
pandas>=2.2.0
numpy>=1.26.0
openpyxl>=3.1.2
plotly>=5.22.0
```

## Run locally

```
pip install -r requirements.txt
streamlit run app.py
```

The app opens at `http://localhost:8501`.

## Deploy on Streamlit Community Cloud

Push this repo to GitHub, then connect it from the Streamlit Cloud dashboard. The app file is `app.py`. No environment variables are required.

## Input file format

The tool expects the p44 unified shipment export in either CSV or Excel form. The minimum required columns are `Shipment ID`, `Current mode`, `Current carrier`, and the Stop 1 family (`Stop 1`, `Stop 1 city`, `Stop 1 state`, `Stop 1 country`, `Stop 1 actual arrival time`, `Stop 1 actual departure time`). Stops 2 through 20 follow the same naming convention.

Only rows with `Current mode == TRUCKLOAD` are kept; AIR, RAIL, OCEAN, and UNKNOWN rows are dropped automatically. Tenant column is optional; if absent, a placeholder tenant name is used. SCAC is not present in the TL export, so the SCAC column in the output is intentionally blank.

## Files in this repo

- `app.py` — the Streamlit application
- `requirements.txt` — Python dependency list
- `DESIGN.md` — design specification and design decisions
- `README.md` — this file

## Notes

Whole Journey mode requires all stop arrival and departure timestamps for the shipment's stop count to be present for the full breakdown; partial timestamps just leave the relevant segments as NaN without dropping the shipment. Negative segment durations (caused by corrupted timestamps) are excluded per-metric only, so one bad timestamp on a shipment doesn't poison all of its other segment statistics.

Default mode runs purely on the picked start/end milestone pair; only those two timestamps need to be present per shipment, and the lane is defined as first stop endpoint to last stop endpoint regardless of how many intermediate stops the shipment has.

All ranking logic — MAD-from-lane-median, percentile deviation, volume, alphabetical — is identical to the ocean tool. The mental model is "least deviated carrier," not "fastest carrier."
