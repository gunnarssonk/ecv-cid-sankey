#!/usr/bin/env python3
"""
Renders the two Sankey diagrams published with the mapping workbook "Mapping
GCOS ECVs and ESA ECV projects to IPCC AR6 Climatic Impact-Drivers".

The workbook is archived separately and can be downloaded from https://doi.org/10.5281/zenodo.21871535.
Either put it in a ``data/`` directory beside this file or pass ``--workbook``.

Functions:
- warn(message): Reports a data problem on stderr without stopping the run.
- resolve_workbook(explicit): Locates the mapping workbook, or explains how to supply it.
- load_flows(workbook, spec): Reads one mapping sheet and returns its non-excluded links.
- apply_abbreviations(flows): Shortens over-long node names and builds the footer note.
- build_nodes(flows, spec): Orders the ECV and CID columns for drawing.
- build_subtitle(flows, spec): Fills the subtitle's counts from the data.
- column_shapes(flows, spec): Measures each column, so one row pitch can fit every figure.
- build_payload(flows, spec, columns): Assembles the nodes, links, text and style for the browser.
- d3_script(): Inlines the vendored plotting library.
- render_html(payload): Fills the HTML template with the data and the library.
- build_figure(spec, workbook, outdir, columns): Reads one sheet, writes one HTML file.
- parse_args(argv): Defines and parses the command-line options.
- main(argv): Measures every figure, builds the requested ones, opens them.

Getting started:
- Put the workbook in ./data/, or pass --workbook PATH.
- Run: python ecv_cid_sankey.py [gcos|esa|both] [--outdir DIR] [--no-open]
- Edit FIGURES to change what a figure says; edit HTML_TEMPLATE to change how it is drawn.
"""

from __future__ import annotations

import argparse
import json
import sys
import webbrowser
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

__version__ = "1.0.0"

HERE = Path(__file__).resolve().parent


WORKBOOK_NAME   = "gcos-ecv-esa-to-ipcc-cid_v1.0.xlsx"
WORKBOOK_DOI    = "https://doi.org/10.5281/zenodo.21871535"
HEADER_ROW      = 4
DIRECT          = 2  # direct input to a CID index, or the CID's primary measurement variable
INDIRECT        = 1  # contributes through a documented mechanism, but does not drive the CID


CATEGORY_COLOURS = {
    "Heat and Cold": "#e6c35a",
    "Wet and Dry":   "#e1a04b",
    "Wind":          "#d78ca5",
    "Snow and Ice":  "#aa96c8",
    "Coastal":       "#78afdc",
    "Open Ocean":    "#82c38c",
    "Other":         "#b9b9be",
}
CATEGORY_ORDER = list(CATEGORY_COLOURS)

ABBREVIATIONS = {
    "Fraction of Absorbed Photosynthetically Active Radiation": "FAPAR",
    "Carbon Dioxide, Methane and Other Greenhouse Gases": "CO₂, CH₄ and other GHGs",
    "Atmospheric Carbon Dioxide at Surface": "Atmospheric CO₂ at Surface",
}

FOOTNOTE_MARKERS = "*†‡"

A4_WIDTH  = 790   # 210 mm = 793.7 px
A4_HEIGHT = 1118  # 297 mm = 1122.5 px
MARGIN_X  = 28    # page margin


# --- Per-figure settings ----------------------------------------------------
@dataclass(frozen=True)
class FigureSpec:
    """Everything that differs between the two figures."""

    key: str
    sheet: str
    domain_column: str
    domain_order: list  # top-to-bottom order of the left column's blocks
    title: str
    subtitle: str  # format string over {n_ecvs}, {n_cids}, {n_links} and {indirect_pct}
    source: str
    left_column_title: str
    output_name: str

    # Only the two label gutters are chosen; the flows take whatever is left of
    # the A4 width, so the page can never come out the wrong size.
    pad_left: int    # ECV label gutter
    pad_right: int   # CID label gutter

    @property
    def flow_width(self) -> int:
        """Width of the sankey curves: the A4 page less the margins and gutters."""
        return A4_WIDTH - 2 * MARGIN_X - self.pad_left - self.pad_right


FIGURES = {
    "gcos": FigureSpec(
        key="gcos",
        sheet="4.2 GCOS ECV-CID Mapping",
        domain_column="ecv_domain",
        domain_order=["Atmosphere", "Land", "Ocean"],
        title=(
            "Mapping GCOS Essential Climate Variables (ECVs) "
            "to IPCC Climatic Impact-Drivers (CIDs)"
        ),
        subtitle=(
            "The {n_ecvs} GCOS ECVs inform {n_cids} IPCC Climatic Impact-Drivers "
            "across {n_links} mapped links; {indirect_pct}% are indirect, "
            "relying on derived products."
        ),
        source="ECVs: GCOS-245 (2022). CIDs: IPCC AR6 WGI Ch.12. Mapping: ESA, doi.org/10.5281/zenodo.21871535",
        left_column_title="ECV DOMAINS",
        output_name="gcos-ecv-cid.html",
        pad_left=202,
        pad_right=166,
    ),
    "esa": FigureSpec(
        key="esa",
        sheet="5.2 ESA ECV-CID Mapping",
        domain_column="esa_cryosphere_domain", # Cryosphere is a domain in the ESA figure, but not in the GCOS figure
        domain_order=["Atmosphere", "Land", "Cryosphere", "Ocean"],
        title=(
            "Mapping ESA CCI Essential Climate Variable (ECV) projects "
            "to IPCC Climatic Impact-Drivers (CIDs)"
        ),

        subtitle=(
            "The {n_ecvs} ESA CCI ECV projects inform {n_cids} of the 33 IPCC "
            "Climatic Impact-Drivers across {n_links} mapped links; "
            "{indirect_pct}% are indirect, relying on derived products."
        ),
        source="ECV projects: ESA CCI. CIDs: IPCC AR6 WGI Ch.12. Mapping: ESA, doi.org/10.5281/zenodo.21871535",
        left_column_title="ESA CCI ECV DOMAINS",
        output_name="esa-ecv-cid.html",
        pad_left=194,
        pad_right=166,
    ),
}


# --- Loading and preparing the mapping --------------------------------------
def warn(message: str) -> None:
    """Report a data problem on stderr without stopping the run.

    Every check here warns rather than raises: a figure that draws with a note
    about one odd row is more useful than no figure at all.
    """
    print(f"warning: {message}", file=sys.stderr)


def resolve_workbook(explicit: Path | None) -> Path:
    """Locate the mapping workbook, or explain how to supply it."""
    if explicit is not None:
        path = explicit.expanduser().resolve()
        if not path.is_file():
            raise SystemExit(f"error: no workbook at {path}")
        return path

    # fromkeys de-duplicates but keeps the search order: running the script from
    # its own directory would otherwise list the same path twice.
    candidates = list(dict.fromkeys([
        HERE / "data" / WORKBOOK_NAME,
        HERE / WORKBOOK_NAME,
        Path.cwd() / "data" / WORKBOOK_NAME,
    ]))
    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()

    searched = "\n".join(f"  {c}" for c in candidates)
    raise SystemExit(
        f"error: could not find {WORKBOOK_NAME}. Looked in:\n{searched}\n"
        f"Download it from {WORKBOOK_DOI} and place it in a 'data' directory "
        f"beside this script, or pass --workbook PATH."
    )


def load_flows(workbook: Path, spec: FigureSpec) -> pd.DataFrame:
    """Read one sheet into ``domain, ecv_source, cid_target, cid_type, value``.
    Excluded rows are dropped and the category names canonicalised.
    """
    text_columns = [spec.domain_column, "ecv_source", "cid_target", "cid_type"]
    flows = pd.read_excel(
        workbook,
        sheet_name=spec.sheet,
        header=HEADER_ROW,
        usecols=text_columns + ["value"],
    )
    flows.columns = flows.columns.str.strip()

    # "string" keeps blanks as NA rather than the text "nan", so a row missing an
    # endpoint is dropped below instead of drawn as a node called "nan".
    for column in text_columns:
        flows[column] = flows[column].astype("string").str.strip()

    incomplete = flows[text_columns].isna().any(axis=1) | (flows[text_columns] == "").any(axis=1)
    if incomplete.any():
        warn(f"{spec.sheet}: dropped {int(incomplete.sum())} row(s) "
             f"with a blank endpoint or category")
        flows = flows[~incomplete]

    values = pd.to_numeric(flows["value"], errors="coerce")
    if values.isna().any():
        warn(f"{spec.sheet}: {int(values.isna().sum())} row(s) have a non-numeric value; treated as excluded")
    flows = flows.assign(value=values.fillna(0))

    unexpected = sorted(float(v) for v in set(flows["value"].unique()) - {0, INDIRECT, DIRECT})
    if unexpected:
        warn(f"{spec.sheet}: unexpected link value(s) {unexpected}; anything but {DIRECT} draws as indirect")

    flows = flows[flows["value"] > 0]  # 0 is assessed-but-excluded, not a link
    if flows.empty:
        raise SystemExit(f"error: {spec.sheet} contains no links with a value above zero")

    unknown = sorted(set(flows["cid_type"]) - set(CATEGORY_ORDER))
    if unknown:
        warn(
            f"{spec.sheet}: CID category {unknown} is not one of the {len(CATEGORY_ORDER)} "
            f"IPCC categories; it will draw in grey below the known ones"
        )

    unknown_domains = sorted(set(flows[spec.domain_column]) - set(spec.domain_order))
    if unknown_domains:
        warn(f"{spec.sheet}: domain {unknown_domains} is not in domain_order; it will sort last")

    flows = flows.rename(columns={spec.domain_column: "domain"})
    return flows[["domain", "ecv_source", "cid_target", "cid_type", "value"]]


def apply_abbreviations(flows: pd.DataFrame) -> tuple[pd.DataFrame, dict, list]:
    """Shorten over-long node names, and build the footer note explaining them."""
    flows = flows.copy()
    present = set()
    for column in ("ecv_source", "cid_target"):  # the columns holding node names
        present |= set(flows[column])            # the long names, before renaming
        flows[column] = flows[column].replace(ABBREVIATIONS)

    # Only one-word acronyms are footnoted; a shortened phrase such as
    # "Atmospheric CO2 at Surface" already reads for itself.
    noted = [
        (short, long)
        for long, short in ABBREVIATIONS.items()
        if long in present and " " not in short
    ]
    if len(noted) > len(FOOTNOTE_MARKERS):
        warn(f"only {len(FOOTNOTE_MARKERS)} footnote markers available for {len(noted)} abbreviations")
        noted = noted[: len(FOOTNOTE_MARKERS)]

    marks = {short: FOOTNOTE_MARKERS[i] for i, (short, _) in enumerate(noted)}
    notes = [f"{marks[short]} {short}: {long}" for short, long in noted]

    # After the renaming, so two names collapsing to one label merge into one link.
    flows = flows.groupby(
        ["domain", "ecv_source", "cid_target", "cid_type"], as_index=False
    )["value"].sum()

    return flows, marks, notes


def build_nodes(flows: pd.DataFrame, spec: FigureSpec) -> tuple[list, list]:
    """Order the two columns: ECVs by domain then name, CIDs by category then name."""
    ranks = {name: i for i, name in enumerate(spec.domain_order)}
    ecv_pairs = (
        flows[["domain", "ecv_source"]]
        .drop_duplicates()
        # fillna(len(ranks)) puts anything unlisted after everything listed.
        .assign(_rank=lambda d: d["domain"].map(ranks).fillna(len(ranks)))
        .sort_values(["_rank", "ecv_source"])
    )
    ecv_nodes = [
        {"name": row.ecv_source, "domain": row.domain} for row in ecv_pairs.itertuples()
    ]
    # A variable filed under two domains gets a row in each, which reads as two
    # different variables -- worth saying out loud rather than drawing twice.
    split = ecv_pairs["ecv_source"].value_counts()
    split = sorted(split[split > 1].index)
    if split:
        warn(f"{split} appear under more than one domain and will be drawn once per domain")

    # A CID reachable through several category labels belongs to whichever
    # category carries the most value into it.
    category_ranks = {name: i for i, name in enumerate(CATEGORY_ORDER)}
    cid_pairs = (
        flows.groupby(["cid_target", "cid_type"], as_index=False)["value"]
        .sum()
        .sort_values("value", ascending=False)
        .drop_duplicates("cid_target")
        .assign(_rank=lambda d: d["cid_type"].map(category_ranks).fillna(len(category_ranks)))
        .sort_values(["_rank", "cid_target"])
    )
    cid_nodes = [
        {"name": row.cid_target, "category": row.cid_type} for row in cid_pairs.itertuples()
    ]
    return ecv_nodes, cid_nodes


def build_subtitle(flows: pd.DataFrame, spec: FigureSpec) -> str:
    """Fill the subtitle's counts from the data."""
    n_links = len(flows)
    n_indirect = n_links - int((flows["value"] == DIRECT).sum())

    return spec.subtitle.format(
        n_links=n_links,
        n_ecvs=flows["ecv_source"].nunique(),
        n_cids=flows["cid_target"].nunique(),
        indirect_pct=round(n_indirect / n_links * 100),
    )


def column_shapes(flows: pd.DataFrame, spec: FigureSpec) -> list:
    """Return ``[rows, gaps]`` for each column. Every figure's shapes go to every
    figure, so one row pitch fits them all and their labels read alike."""
    flows, _, _ = apply_abbreviations(flows)
    ecv_nodes, cid_nodes = build_nodes(flows, spec)
    return [
        [len(ecv_nodes), len({n["domain"] for n in ecv_nodes}) - 1],
        [len(cid_nodes), len({n["category"] for n in cid_nodes}) - 1],
    ]


def build_payload(flows: pd.DataFrame, spec: FigureSpec, columns: list) -> dict:
    """Assemble everything the browser needs: nodes, links, text and style."""
    flows, marks, notes = apply_abbreviations(flows)
    ecv_nodes, cid_nodes = build_nodes(flows, spec)

    links = [
        {
            "source": row.ecv_source,
            "target": row.cid_target,
            "category": row.cid_type,
            "direct": bool(row.value == DIRECT),
        }
        for row in flows.itertuples()
    ]

    return {
        "links": links,
        "ecv_nodes": ecv_nodes,
        "cid_nodes": cid_nodes,
        "category_order": [c for c in CATEGORY_ORDER if c in set(flows["cid_type"])],
        "colors": CATEGORY_COLOURS,
        "abbrev_marks": marks,
        "abbrev_notes": notes,
        "title": spec.title,
        "subtitle": build_subtitle(flows, spec),
        "source": spec.source,
        "left_column_title": spec.left_column_title,
        "style": {
            "pad_left": spec.pad_left,
            "pad_right": spec.pad_right,
            "flow_width": spec.flow_width,
            "margin_x": MARGIN_X,
            "a4_width": A4_WIDTH,
            "a4_height": A4_HEIGHT,
            "columns": columns,  # every column of every figure; see column_shapes
        },
    }


# --- Sankeys ----------------------------------------------------------------
# render_html fills __DATA_JSON__, __D3_SCRIPT__ and __PAGE_TITLE__.

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__PAGE_TITLE__</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: 'Segoe UI', sans-serif;
         background: #f2f2f0; color: #333; padding: 28px;
         display: flex; justify-content: center; }
  #chart { background: #fff; display: inline-block;
           box-shadow: 0 1px 4px rgba(0,0,0,0.10); }
  @media print {
    @page { size: A4 portrait; margin: 0; }
    body { background: #fff; padding: 0; }
    #chart { box-shadow: none; }
    .tooltip { display: none; }
  }
  svg text { font-family: 'Segoe UI', sans-serif; }
  .link { fill: none; transition: stroke-opacity 0.15s; }
  .link:hover { stroke-opacity: 0.95; }
  .tooltip { position: fixed; background: rgba(30,30,30,0.92); color: #eee;
    border-radius: 4px; padding: 6px 10px; font-size: 12px; pointer-events: none;
    opacity: 0; transition: opacity 0.12s; z-index: 100; line-height: 1.45; }
</style>
</head>
<body>
<div id="chart"></div>
<div id="tooltip" class="tooltip"></div>
__D3_SCRIPT__
<script>

const DATA = __DATA_JSON__;
const STYLE = DATA.style;
const tooltip = document.getElementById("tooltip");


// ---- COLOUR --------------------------------------------------------------
const COLORS = DATA.colors;
const LINK_OPACITY = 0.50;  // the flow curves and legend samples
const MARK_OPACITY = 0.75;  // the CID bars and the legend swatches
const BAND_TINT    = 0.80;  // category bands washed towards white (1 = white)

const SECTION_TITLE = "#4d4d4d";     // top column titles (ECV DOMAINS / CID CATEGORIES)
const DOMAIN_SPINE = "#b4b4b4";      // thin bar behind each domain block
const BAND         = "#f5f6f7";      // grey box behind each domain's ECVs
const INK = "#2b2b2b", MUTED = "#6b6b6b", RULE = "#d9dcdf";


// ---- TYPE -----------------------------------------------------------------
const TITLE_SIZE = 12, SUB_SIZE = 10, NOTE_SIZE = 9;
const ECV_SIZE = 9.5, CID_SIZE = 9;                 // node labels
const HEADER_SIZE = 9.5, HEADER_TRACK = 0.1;        // block headers
const TOPTITLE_SIZE = 8.5, TOPTITLE_TRACK = 1.4;    // column titles, CAPS


// ---- FLOWS ----------------------------------------------------------------
// A bar always equals the sum of its links.
const DIRECT_W = 1.62;
const INDIRECT_W = 0.81;
const INDIRECT_DASH = "3,2.4";
const STUB = 10;         // flat run at the ECV end before the curve starts


// ---- COLUMNS ---------------------------------------------------------------
const ROW_MAX = 15;      // loosest and tightest a row may be set
const ROW_MIN = 12;
const BLOCK_GAP = 36;
const NODE_GAP  = 2;     // smallest blank gap allowed btw bars
const BAR_W = 2.5;       // CID bar width - same as the ECV domains
const BAND_GAP = 3;      // btw box and node
const BAND_PAD = 4;      // btw boxes


// ---- PAGE ------------------------------------------------------------------
// WIDTH comes out of the gutters and must equal A4_W.
// HEIGHT falls out of the row count and checked against A4_H at the end.
const A4_W = STYLE.a4_width, A4_H = STYLE.a4_height;
const M = { top: 0, right: STYLE.margin_x, left: STYLE.margin_x };  // top set below
const PAD_L  = STYLE.pad_left;    // ECV label gutter
const PAD_R  = STYLE.pad_right;   // CID label gutter
const FLOW_W = STYLE.flow_width;  // sankey flow-curve area
const WIDTH  = M.left + PAD_L + FLOW_W + PAD_R + M.right;

// Each step is measured off the one before it, so changing a gap moves everything below it.
const TITLE_Y = 34;          // btw page top and title
const SUB_GAP = 19;          // btw title, subtitle, legend
const HEAD_RULE_GAP = 12;    // btw subtitle and under header
const LEGEND_ROW_GAP = 16;   // btw CID and connection row
const HEAD_PAD = 45;         // btw legend and column titles
const TOPTITLE_GAP = 18;     // btw column title and block header
const HEADER_PAD = 13;       // btw block header and its block
const RULE_GAP = 16;         // columns, above the source line
const NOTE_GAP = 20;         // source line
const FOOT_PAD = 20;         // bottom

const SUB_Y0      = TITLE_Y + SUB_GAP;
const HEAD_RULE_Y = SUB_Y0 + HEAD_RULE_GAP;
const LEGEND_Y    = HEAD_RULE_Y + SUB_GAP;
M.top = LEGEND_Y + LEGEND_ROW_GAP + HEAD_PAD + TOPTITLE_GAP;

const CHROME = M.top + RULE_GAP + NOTE_GAP + FOOT_PAD;  // page less the columns


// ---- NODES AND LINKS -------------------------------------------------------
const ecvs = DATA.ecv_nodes.map((d, i) => ({ ...d, i, h: 0 }));
const cids = DATA.cid_nodes.map((d, i) => ({ ...d, i, h: 0 }));
const ecvByName = Object.fromEntries(ecvs.map(d => [d.name, d]));
const cidByName = Object.fromEntries(cids.map(d => [d.name, d]));

const links = DATA.links
  .filter(l => ecvByName[l.source] && cidByName[l.target])
  .map(l => ({
    ...l,
    si: ecvByName[l.source].i,
    ti: cidByName[l.target].i,
    w: l.direct ? DIRECT_W : INDIRECT_W,
  }));

links.forEach(l => { ecvs[l.si].h += l.w; cids[l.ti].h += l.w; });
[...ecvs, ...cids].forEach(d => { d.h = Math.max(1.5, d.h); });


// ---- BLOCKS ----------------------------------------------------------------
//  Each category is one continuous run of rows. These are what the bands and headers draw.
const runs = (items, key) => items.reduce((acc, d) => {
  const last = acc[acc.length - 1];
  if (last && last.name === key(d)) last.items.push(d);
  else acc.push({ name: key(d), items: [d] });
  return acc;
}, []);

const blocks = runs(ecvs, d => d.domain);    // ECVs, by domain
const groups = runs(cids, d => d.category);  // CIDs, by category


// ---- ROW PITCH -------------------------------------------------------------
const tallestRow = ([rows, gaps]) => (A4_H - CHROME - gaps * BLOCK_GAP) / rows;
const ROW = Math.max(ROW_MIN, Math.min(ROW_MAX, ...STYLE.columns.map(tallestRow)));


// ---- PLACE ROWS ------------------------------------------------------------
function place(runs) {
  let y = 0;
  runs.forEach((r, i) => {
    if (i) y += BLOCK_GAP;
    r.top = y;
    r.items.forEach((d, j) => {
      d.cy = y + ROW * (j + 0.5);
      d.y0 = d.cy - d.h / 2;
    });
    r.bot = y += ROW * r.items.length;
  });
  return y;
}

// Shorter column ends higher.
const colH = Math.max(place(blocks), place(groups));


// ---- LINK ENDS -------------------------------------------------------------
function stackEnds(nodes, near, far, y) {
  nodes.forEach(d => {
    let c = d.y0;
    links.filter(l => l[near] === d.i)
         .sort((a, b) => a[far] - b[far])
         .forEach(l => { l[y] = c + l.w / 2; c += l.w; });
  });
}
stackEnds(ecvs, "si", "ti", "sy");
stackEnds(cids, "ti", "si", "ty");

const linkPath = l => {
  const xm = (STUB + FLOW_W) / 2;
  return `M0,${l.sy}L${STUB},${l.sy}C${xm},${l.sy} ${xm},${l.ty} ${FLOW_W},${l.ty}`;
};


// Darkens (amt < 0) or lightens (amt > 0) a hex colour.
function shade(hex, amt) {
  const n = parseInt(hex.slice(1), 16);
  const c = [n >> 16 & 255, n >> 8 & 255, n & 255].map(v =>
    Math.max(0, Math.min(255, Math.round(v + (amt < 0 ? -v : 255 - v) * Math.abs(amt))))
  );
  return `rgb(${c[0]},${c[1]},${c[2]})`;
}

// ---- CANVAS ----------------------------------------------------------------
const svg = d3.select("#chart").append("svg")
  .attr("width", WIDTH).attr("xmlns", "http://www.w3.org/2000/svg");
const bg = svg.append("rect").attr("width", WIDTH).attr("fill", "#fff");

// Inside g, x = 0 is where the flows start and y = 0 is the first row; negative
// x is the left gutter. Page furniture is drawn on svg instead, off X0.
const g = svg.append("g")
  .attr("transform", `translate(${M.left + PAD_L}, ${M.top})`);


const BAND_L_X0 = -PAD_L, BAND_L_X1 = -(BAND_GAP + BAR_W + 2.5);
const BAND_R_X0 = FLOW_W + BAR_W + BAND_GAP, BAND_R_X1 = FLOW_W + PAD_R;
const BAND_L_MID = (BAND_L_X0 + BAND_L_X1) / 2;
const BAND_R_MID = (BAND_R_X0 + BAND_R_X1) / 2;
const LABEL_L_X = -12, LABEL_R_X = FLOW_W + BAR_W + 6;


// ---- DRAW: BACKDROP --------------------------------------------------------
// grey domain bands behind the ECVs
g.selectAll("rect.band").data(blocks).join("rect").attr("class", "band")
  .attr("x", BAND_L_X0).attr("y", d => d.top - BAND_PAD)
  .attr("width", BAND_L_X1 - BAND_L_X0).attr("height", d => d.bot - d.top + BAND_PAD * 2)
  .attr("fill", BAND);

g.selectAll("rect.spine").data(blocks).join("rect").attr("class", "spine")
  .attr("x", -5).attr("y", d => d.top - BAND_PAD)
  .attr("width", BAR_W).attr("height", d => d.bot - d.top + BAND_PAD * 2)
  .attr("fill", DOMAIN_SPINE);

// coloured bands behind the CIDs, one per category, tinted from COLORS and offset past the bars by BAND_GAP
g.selectAll("rect.cidband").data(groups).join("rect").attr("class", "cidband")
  .attr("x", BAND_R_X0).attr("y", d => d.top - BAND_PAD)
  .attr("width", BAND_R_X1 - BAND_R_X0)
  .attr("height", d => d.bot - d.top + BAND_PAD * 2)
  .attr("fill", d => shade(COLORS[d.name] || "#999", BAND_TINT));

// ---- DRAW: FLOWS -----------------------------------------------------------
g.append("g").selectAll("path").data(links).join("path")
  .attr("class", "link").attr("d", linkPath)
  .attr("stroke", d => COLORS[d.category] || "#999")
  .attr("stroke-opacity", LINK_OPACITY)
  .attr("stroke-width", d => d.w)
  .attr("stroke-dasharray", d => d.direct ? null : INDIRECT_DASH)
  .on("mouseenter", (e, d) => {
    tooltip.style.opacity = 1;
    tooltip.innerHTML = `<strong>${d.source}</strong> &rarr; <strong>${d.target}</strong>` +
      `<br>${d.category} · ${d.direct ? "Direct" : "Indirect"}`;
  })
  .on("mousemove", e => {
    tooltip.style.left = (e.clientX + 14) + "px";
    tooltip.style.top  = (e.clientY - 10) + "px";
  })
  .on("mouseleave", () => { tooltip.style.opacity = 0; });

// ---- DRAW: MARKS -----------------------------------------------------------
// The CID bars, drawn over the flows they gather.
g.append("g").selectAll("rect.cid").data(cids).join("rect").attr("class", "cid")
  .attr("x", FLOW_W).attr("y", d => d.y0)
  .attr("width", BAR_W).attr("height", d => d.h)
  .attr("fill", d => COLORS[d.category] || "#666")
  .attr("fill-opacity", MARK_OPACITY);

// ---- DRAW: TEXT ------------------------------------------------------------
// Footnote markers, appended to the drawn label only - the node name stays
const MARKS = DATA.abbrev_marks || {};
const marked = d => d.name + (MARKS[d.name] || "");

// left column: domain headers, its title, then the ECV labels
g.selectAll("text.dom").data(blocks).join("text").attr("class", "dom")
  .attr("x", LABEL_L_X).attr("y", d => d.top - HEADER_PAD).attr("text-anchor", "end")
  .attr("font-size", HEADER_SIZE).attr("font-weight", 700).attr("letter-spacing", HEADER_TRACK)
  .attr("fill", SECTION_TITLE)  // same grey as the column title, so the two read as one hierarchy
  .text(d => d.name + "  (" + d.items.length + ")");

g.append("text").attr("class", "coltitle")
  .attr("x", BAND_L_MID).attr("y", blocks[0].top - HEADER_PAD - TOPTITLE_GAP)
  .attr("text-anchor", "middle")
  .attr("font-size", TOPTITLE_SIZE).attr("font-weight", 700).attr("letter-spacing", TOPTITLE_TRACK)
  .attr("fill", SECTION_TITLE)
  .text(DATA.left_column_title);

g.append("g").selectAll("text.ecv").data(ecvs).join("text").attr("class", "ecv")
  .attr("x", LABEL_L_X).attr("y", d => d.cy).attr("dy", "0.34em")
  .attr("text-anchor", "end").attr("font-size", ECV_SIZE).attr("fill", INK)
  .text(marked);

// right column: the same, mirrored
g.selectAll("text.grouphead").data(groups).join("text").attr("class", "grouphead")
  .attr("x", LABEL_R_X).attr("y", d => d.top - HEADER_PAD)
  .attr("font-size", HEADER_SIZE).attr("font-weight", 700).attr("letter-spacing", HEADER_TRACK)
  .attr("fill", d => COLORS[d.name] ? shade(COLORS[d.name], -0.35) : INK)
  .text(d => d.name + "  (" + d.items.length + ")");

g.append("text").attr("class", "coltitle")
  .attr("x", BAND_R_MID).attr("y", groups[0].top - HEADER_PAD - TOPTITLE_GAP)
  .attr("text-anchor", "middle")
  .attr("font-size", TOPTITLE_SIZE).attr("font-weight", 700).attr("letter-spacing", TOPTITLE_TRACK)
  .attr("fill", SECTION_TITLE)
  .text("CID CATEGORIES");

g.append("g").selectAll("text.cidlabel").data(cids).join("text").attr("class", "cidlabel")
  .attr("x", LABEL_R_X).attr("y", d => d.cy).attr("dy", "0.34em")
  .attr("font-size", CID_SIZE).attr("fill", INK)
  .text(marked);


// ---- DRAW: HEADER ----------------------------------------------------------
// The title, subtitle and source line are written on the Python side.
const X0 = M.left;

svg.append("text").attr("x", X0).attr("y", TITLE_Y)
  .attr("font-size", TITLE_SIZE).attr("font-weight", 700).attr("fill", INK)
  .text(DATA.title);

svg.append("text").attr("x", X0).attr("y", SUB_Y0)
  .attr("font-size", SUB_SIZE).attr("fill", MUTED).text(DATA.subtitle);

svg.append("line").attr("x1", X0).attr("x2", WIDTH - M.right)
  .attr("y1", HEAD_RULE_Y).attr("y2", HEAD_RULE_Y).attr("stroke", RULE);


// ---- LEGEND --------------------------------------------------------------
const LEGEND_SIZE      = 9;    // one size for both the row labels and the items
const LEGEND_ICON_GAP  = 5;    // btw icon (swatch/line) and its label
const LEGEND_LABEL_GAP = 8;    // btw a row's label and its first item
const LEGEND_ITEM_GAP  = 16;   // between items within the same row
const SWATCH_W = 9, LINE_W = 16;

function legendRow(label, items, y) {
  const layer = svg.append("g").attr("transform", `translate(${X0}, ${y})`);
  const entries = [{ kind: "label", text: label }, ...items];
  let lx = 0;
  entries.forEach((e, i) => {
    if (i === 1) lx += LEGEND_LABEL_GAP;   // label -> its first item
    else if (i > 1) lx += LEGEND_ITEM_GAP; // between items
    const item = layer.append("g").attr("transform", `translate(${lx}, 0)`);
    let iconW = 0;
    if (e.kind === "swatch") {
      item.append("rect").attr("y", -8).attr("width", SWATCH_W).attr("height", 9)
        .attr("fill", e.color).attr("fill-opacity", MARK_OPACITY);
      iconW = SWATCH_W + LEGEND_ICON_GAP;
    } else if (e.kind === "line") {
      // Same weight, dash and opacity as the flows; only the hue is
      // neutralised, since the row above already explains the colour.
      item.append("line").attr("x1", 0).attr("x2", LINE_W).attr("y1", -3).attr("y2", -3)
        .attr("stroke", INK).attr("stroke-width", e.width)
        .attr("stroke-opacity", LINK_OPACITY)
        .attr("stroke-dasharray", e.dash || null);
      iconW = LINE_W + LEGEND_ICON_GAP;
    }
    // Row labels differ from their items in weight alone.
    const textEl = item.append("text").attr("x", iconW)
      .attr("font-size", LEGEND_SIZE).attr("fill", INK)
      .attr("font-weight", e.kind === "label" ? 700 : null)
      .text(e.text);
    lx += iconW + textEl.node().getComputedTextLength();
  });
}

// The legend encodes the CID category only
legendRow("CID (color):",
  DATA.category_order.map(cat => ({ kind: "swatch", color: COLORS[cat], text: cat })), LEGEND_Y);
legendRow("Connection (line):", [
  { kind: "line", width: DIRECT_W, text: "Direct" },
  { kind: "line", width: INDIRECT_W, text: "Indirect", dash: INDIRECT_DASH },
], LEGEND_Y + LEGEND_ROW_GAP);

// ---- DRAW: FOOTER ----------------------------------------------------------
const ruleY = M.top + colH + RULE_GAP;
svg.append("line").attr("x1", X0).attr("x2", WIDTH - M.right)
  .attr("y1", ruleY).attr("y2", ruleY).attr("stroke", RULE);

const sourceY = ruleY + NOTE_GAP;
svg.append("text").attr("x", X0).attr("y", sourceY)
  .attr("font-size", NOTE_SIZE).attr("fill", MUTED)
  .attr("font-style", "italic").text(DATA.source);

if ((DATA.abbrev_notes || []).length) {
  svg.append("text").attr("x", WIDTH - M.right).attr("y", sourceY)
    .attr("text-anchor", "end")
    .attr("font-size", NOTE_SIZE).attr("fill", MUTED)
    .attr("font-style", "italic")
    .text(DATA.abbrev_notes.join("; "));
}

const HEIGHT = sourceY + FOOT_PAD;
svg.attr("height", HEIGHT).attr("viewBox", `0 0 ${WIDTH} ${HEIGHT}`);
bg.attr("height", HEIGHT);

// Height follows row count, e.g. more ECVs / CIDs push the figure onto a second sheet.

console.log(`artboard ${WIDTH} x ${Math.round(HEIGHT)}px (A4 budget ${A4_W} x ${A4_H})`);
if (HEIGHT > A4_H) {
  console.warn(`${(HEIGHT - A4_H).toFixed(0)}px too tall for one A4 page. Rows are ` +
    `already at ROW_MIN (${ROW_MIN}px) and cannot tighten further -- lower ROW_MIN ` +
    `to fit, at the cost of crowding the labels.`);
}
if (WIDTH > A4_W) {
  console.warn(`${(WIDTH - A4_W).toFixed(0)}px too wide for one A4 page. ` +
    `Narrow pad_left / pad_right for this figure in FIGURES.`);
}

const tallest = Math.max(...ecvs.map(d => d.h), ...cids.map(d => d.h));
if (tallest > ROW - NODE_GAP) {
  console.warn(`The heaviest bar is ${tallest.toFixed(1)}px in a ${ROW}px row, so ` +
    `neighbouring bars come within ${(ROW - tallest).toFixed(1)}px of touching. ` +
    `Scale DIRECT_W/INDIRECT_W down to restore the ${NODE_GAP}px gap.`);
}

</script>
</body>
</html>
"""

D3_VENDOR_PATH = HERE / "vendor" / "d3-selection-3.0.0.min.js"


def d3_script() -> str:
    """Inline the vendored plotting library, so the figure renders offline."""
    if not D3_VENDOR_PATH.is_file():
        raise SystemExit(
            f"error: no plotting library at {D3_VENDOR_PATH}. Restore the "
            f"'vendor' directory from the repository alongside this script."
        )
    return f"<script>\n{D3_VENDOR_PATH.read_text(encoding='utf-8').strip()}\n</script>"


def render_html(payload: dict) -> str:
    """Fill the template with the figure's data and its plotting library."""
    # "</" is escaped so no label in the data can close the <script> early.
    data_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
    page_title = (
        payload["title"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    )
    return (
        HTML_TEMPLATE.replace("__D3_SCRIPT__", d3_script())
        .replace("__PAGE_TITLE__", page_title)
        .replace("__DATA_JSON__", data_json)
    )


def build_figure(spec: FigureSpec, workbook: Path, outdir: Path, columns: list) -> Path:
    """Read one sheet, draw one figure, write one HTML file. Returns its path."""
    flows = load_flows(workbook, spec)
    payload = build_payload(flows, spec, columns)

    outdir.mkdir(parents=True, exist_ok=True)
    output_path = outdir / spec.output_name
    output_path.write_text(render_html(payload), encoding="utf-8")

    n_direct = sum(1 for link in payload["links"] if link["direct"])
    n_links = len(payload["links"])
    print(
        f"{spec.key}: {len(payload['ecv_nodes'])} ECVs, {len(payload['cid_nodes'])} CIDs, "
        f"{n_links} links ({n_direct} direct, {n_links - n_direct} indirect)"
    )
    print(f"{spec.key}: wrote {output_path}")
    return output_path


# --- Command line -----------------------------------------------------------
def parse_args(argv: list | None = None) -> argparse.Namespace:
    """Define and parse the command-line options."""
    parser = argparse.ArgumentParser(
        description=__doc__.split("\n\n")[0],
        epilog=f"Mapping workbook: {WORKBOOK_DOI}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "figure",
        nargs="?",
        default="both",
        choices=[*FIGURES, "both"],
        help="which figure to draw (default: both)",
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        help=f"path to {WORKBOOK_NAME} (default: search ./data and the script's directory)",
    )
    parser.add_argument(
        "--outdir",
        type=Path,
        default=Path.cwd(),
        help="directory to write the HTML into (default: the current directory)",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="write the files without opening them in a browser",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser.parse_args(argv)


def main(argv: list | None = None) -> int:
    """Measure every figure, build the requested ones, open them."""
    args = parse_args(argv)
    workbook = resolve_workbook(args.workbook)
    print(f"reading {workbook}")

    columns = [
        shape
        for spec in FIGURES.values()
        for shape in column_shapes(load_flows(workbook, spec), spec)
    ]

    keys = list(FIGURES) if args.figure == "both" else [args.figure]
    written = [
        build_figure(FIGURES[key], workbook, args.outdir.expanduser().resolve(), columns)
        for key in keys
    ]

    if not args.no_open:
        for path in written:
            try:
                webbrowser.open_new_tab(path.as_uri())
            except webbrowser.Error:
                pass  # headless -- the files are written either way
    return 0


if __name__ == "__main__":
    sys.exit(main())
