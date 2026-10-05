# Sea-Cat Legacy Workbook — Sheet Summary

Source folder: `C:\Users\beethes.ONID\sea-cat\SMC\SMCCopy`

## Files

| File | Size | Role | Readable? |
|---|---|---|---|
| `scdpp.xls` | 195 MB | The **actual database** of record (14 data sheets, ~270k records) | yes |
| `log.xls` | 108 KB | PACER logger registry & processing tracker (2 sheets) | yes |
| `SC.console.xls` | 2.4 MB | Smart Coast data-entry tool (UserForms + logic; its own private sheets) | **encrypted** (RC4, open-password unknown; VBA readable, data not) |
| `ERDPP.console.xls` | 1.0 MB | ERDPP data-entry tool | **encrypted** (same) |

The two large data-entry books are locked; the *data* they edit lives in `scdpp.xls` / `log.xls`, which are fully readable. (`scdpp.xls` was originally too big for Excel's 65 536-row limit and the multibeam inventory is split across five sheets.)

## Universal layout convention

Every sheet uses a three-row header block, then data rows from row 4:

* Row 1 — column captions (print labels, multi-line via `\n`)
* Row 2 — internal field names (e.g. `sc_seamount_id`)
* Row 3 — **data types that double as validation rules**:

| Type | Meaning |
|---|---|
| `id` | primary key: integral number, unique, required |
| `long` / `nmb` | numeric column |
| `date` | Excel serial date |
| `year` | 4-digit year |
| `varchar2(n)` / `char(n)` | text up to `n` characters |
| `list(n)` | `:tok1:tok2:`-delimited list (max `n` chars) |
| `skip` | no validation |

## scdpp.xls — sheets

### SC_* family (Smart Coast record-keeping)

| Sheet | Records | Cols | Contents |
|---|---|---|---|
| `SC_regions` | 278 | 42 | Oceanographic regions: `sc_region_id`, `region_index`, `region_name`, `region_level`, `region_class` (Seamount Group, Archipelago…), `ocean_name`, `plate_name`, bounding box + `REGION_AREA`, `creation_status`, `folder`, `status`. Many rows still `Temporary` / `Error: Zero area`. |
| `SC_seamounts` | 2101 | 101 | **Seamount master table**: `sc_seamount_id`, index names like `SMNT-459S-0001E`, coordinates `lat1..lon2` box, `SEAMOUNT_AREA`, `seamount_name`, alternate names, `folder` (e.g. `GR12`, `WALV`, `LOU`), status flags, grid/map bookkeeping fields. |
| `SC_maps` | 10801 | 68 | **Every produced map job**: `sc_map_id`, `creation_status`, `creation_type` (e.g. `2D:1:MG`, `2D:2:MB`, `2D:3:SS:C`), `folder`, file name, `seamount_index`, `region_index`, long captions, `creation_date`, and GMT/grid parameters (`grid_dx`, `grid_dy`, `margin_y`, `pdf_dpi`, `gmt_normalization`, …). |

### MGG_* family (global grids / multibeam / rock samples)

| Sheet | Records | Cols | Contents |
|---|---|---|---|
| `MGG_multibeam1` | 65 018 | 92 | Multibeam survey inventory (slab 1). `mgg_multibeam_id`, box `lat3/lon3..lat4/lon4`, `folder` (cruise), file name, `multibeam_format_id`, `weight`, `hold`, `released`, `jday1/2` (day-of-year + timestamp), `created info` flags. |
| `MGG_multibeam2` | 64 561 | 86 | Slab 2 |
| `MGG_multibeam3` | 63 569 | 87 | Slab 3 (largest block of orphan `jday` annotations) |
| `MGG_multibeam4` | 186 | 86 | Slab 4 |
| `MGG_multibeam5` | 0 | 85 | Empty template slab |
| `MGG_grids` | 919 | 58 | Gridded datasets (SRTM topography etc.): `mgg_grid_id`, box, `folder` (`SRTM4`…), file name, caption, creation date. |
| `MGG_samples` | 2381 | 28 | **Rock-sample catalogue**: `mgg_sample_id`, box, `er_expedition_name` (link to expeditions), `sample_number` (`MV1203-D06`), `sample_method` (Dredge/Grab/…), label, begin/end lat/lon. |
| `MGG_features` | 4851 | 19 | Named seafloor features: `mgg_feature_id`, `lat0/lon0`, `feature_name`, `region_name`, `ocean_name`, `feature_level`, `feature_type` (Peak / Volcanic Base / Tectonic Basin / …), `feature_elevation`, azimuth/alignment, source. |
| `MGG_expeditions` | 92 | 50 | **Cruise/expedition master**: `er_expedition_id`, `er_expedition_name` (`AII2L25`, `MV1203`, …), name alternatives, upload/duplicate flags, `mdt_code`, ship, leg, NGDC number, sponsor, location. |

### ER_* family (EarthRef / ERDPP bibliographic)

| Sheet | Records | Cols | Contents |
|---|---|---|---|
| `ER_citations` | 157 | 44 | Citations: `er_citation_id`, citation type, long/short authors, `year`, `doi`, title, journal, volume, pages, book title/editors. |
| `ER_expertlevels` | 9 | 7 | Educational expertise levels: `er_level_id`, school-level description, general description, `pacer_log_ids` (logger list), inserted/updated, notes. |

## log.xls — sheets

| Sheet | Records | Cols | Contents |
|---|---|---|---|
| `PACER_loggers` | 32 | 8 | Named data loggers: `pacer_log_id`, first/last name, institution, `pacer_log_ids` (their ID in the `:n:` encoding), inserted/updated dates, remarks. |
| `PACER_logsheet` | 330 | 15 | One row per citation being tracked: `er_citation_id`, online-found / copy-filed / scans / OCR flags, date finished / checked, DB entry + database, `problem`, `pacer_log_ids`, inserted / updated, notes. |

## Hierarchy (how the sheets relate)

```
SC_maps.seamount_index ─────► SC_seamounts.seamount_index
SC_maps.region_index  ─────► SC_regions.region_index
MGG_samples.er_expedition_name ──► MGG_expeditions.er_expedition_name
ER_expertlevels.pacer_log_ids  ─┐
                               ├──► log.xls!PACER_loggers.pacer_log_id
PACER_logsheet.pacer_log_ids  ─┘      (encoded as :n: lists)
SC_* / MGG_* coordinate boxes → SEAMOUNT/REGION/MULTIBEAM_AREA columns,
    with "Error: Zero area ..." marker written into creation_status
```

## Validation findings (from `validate_data.py`, full run)

* `warn` **58 012** — MGG_multibeam3 `jday1/2` cells hold day-of-year numbers with an ISO timestamp annotation (`"23 (2017-01-23T15:04:21.671000)"`); the tool reads only the leading number.
* `ref` **3 263** — orphan references: 1886 samples whose expedition is not in `MGG_expeditions` (**253 distinct cruises**, e.g. `AMAT02RR`, `MV1203`, `HURL`, `KM2201`); 708 maps whose `seamount_index` is not in `SC_seamounts`; 669 maps whose `region_index` is not in `SC_regions`.
* `type` **314** — hard type failures: multibeam3/4 `date1/date2` stored as text dates (139+139+…); `SC_maps` `margin_y`/`pdf_dpi`/`gmt_normalization` length violations; `MGG_grids` file name longer than `varchar2(30)`.
* `lon` **114** — longitudes outside the expected `0..360` range (mostly MGG_multibeam).
* `duplicate` **3** — duplicate primary keys found (`ER_citations` id 2134, one `MGG_samples` id, …).
* `area` **2** — box area recomputed ≠ 0 but the row carries a stale `"Error: Zero area"` marker.

Run it yourself:

```
python validate_data.py                # reads scdpp.xls + log.xls
python validate_data.py PATH.XLS ...   # extra/other workbooks
python validate_data.py --report out.txt
```

Requires `pip install xlrd`. To extend validation to the encrypted `SC.console.xls` / `ERDPP.console.xls`, the workbook open-password is required (`msoffcrypto-tool` is already installed to decrypt). The VBA macros use `boris` for structure/macro protection but that is not the file-open password.