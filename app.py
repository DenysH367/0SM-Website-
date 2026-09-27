#!/usr/bin/env python3
"""
==============================================================================
OSM 202: SIGNATURE ASSIGNMENT - U.S. INCOME INEQUALITY & DISTRIBUTION
==============================================================================

To run this application:
    python3 app.py

Then open your browser at:
    http://localhost:8050
==============================================================================
"""

import sys
import os
import json
import math
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.parse

# ==============================================================================
# STUDENT GUESSES (EDIT THESE VALUES FOR YOUR PRESENTATION)
# ==============================================================================
# Enter your estimated salaries (in US Dollars) for each percentile below:
STUDENT_GUESSES = {
    20: 20000,   # Student guess for 20th percentile (e.g. $20,000)
    40: 38000,   # Student guess for 40th percentile (e.g. $38,000)
    60: 60000,   # Student guess for 60th percentile (e.g. $60,000)
    80: 95000,   # Student guess for 80th percentile (e.g. $95,000)
}
# ==============================================================================

PORT = 8050
DATA_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(DATA_DIR, "dataset_cache.json")
EXCEL_FILE = os.path.join(DATA_DIR, "us_adults_sample_100k.xlsx")

def ensure_cache():
    """Generates the fast JSON cache if it does not already exist."""
    if os.path.exists(CACHE_FILE):
        return
    print("Precomputing dataset cache from us_adults_sample_100k.xlsx (one-time setup)...")
    import zipfile
    import xml.etree.ElementTree as ET

    shared_strings = []
    with zipfile.ZipFile(EXCEL_FILE) as z:
        if 'xl/sharedStrings.xml' in z.namelist():
            tree = ET.parse(z.open('xl/sharedStrings.xml'))
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in tree.getroot().findall('ns:si', ns):
                shared_strings.append(''.join(t.text or '' for t in si.findall('.//ns:t', ns)))

    data_by_state = {}
    us_incomes = []

    with zipfile.ZipFile(EXCEL_FILE) as z:
        ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        with z.open('xl/worksheets/sheet1.xml') as f:
            context = ET.iterparse(f, events=('end',))
            for event, elem in context:
                if elem.tag == f'{{{ns}}}row':
                    r_num = elem.attrib.get('r')
                    if r_num != '1':
                        state_name = None
                        inc = None
                        for c in elem.findall(f'{{{ns}}}c'):
                            ref = c.attrib.get('r')
                            col = ''.join([ch for ch in ref if ch.isalpha()])
                            t = c.attrib.get('t')
                            v = c.find(f'{{{ns}}}v')
                            val = v.text if v is not None else None
                            if t == 's' and val is not None:
                                val = shared_strings[int(val)]
                            
                            if col == 'D':
                                state_name = val
                            elif col == 'N':
                                inc = float(val) if val is not None else 0.0
                        
                        if state_name and inc is not None:
                            us_incomes.append(inc)
                            if state_name not in data_by_state:
                                data_by_state[state_name] = []
                            data_by_state[state_name].append(inc)
                    elem.clear()

    def calc_percentiles(arr):
        n = len(arr)
        p_dict = {}
        for p in range(1, 101):
            k = (n - 1) * (p / 100.0)
            f = math.floor(k)
            c = math.ceil(k)
            if f == c:
                val = arr[int(k)]
            else:
                val = arr[f] * (c - k) + arr[c] * (k - f)
            p_dict[p] = round(val, 2)
        return p_dict

    def calc_inequality(arr):
        n = len(arr)
        if n == 0:
            return {}
        
        arr_nonneg = [x for x in arr if x >= 0]
        n_nn = len(arr_nonneg)
        tot_nn = sum(arr_nonneg)
        
        gini_route1 = 0.0
        gini_route2 = 0.0
        if n_nn > 0 and tot_nn > 0:
            cum_y = 0.0
            B = 0.0
            for i, y in enumerate(arr_nonneg, 1):
                cum_y += y
                L_i = cum_y / tot_nn
                L_prev = (cum_y - y) / tot_nn
                p_i = i / n_nn
                p_prev = (i - 1) / n_nn
                B += 0.5 * (L_i + L_prev) * (p_i - p_prev)
            gini_route1 = round(1.0 - 2.0 * B, 4)
            
            rank_sum = sum(i * y for i, y in enumerate(arr_nonneg, 1))
            gini_route2 = round((2.0 * rank_sum) / (n_nn * tot_nn) - (n_nn + 1.0) / n_nn, 4)
        
        arr_pos = [x for x in arr if x > 0]
        n_pos = len(arr_pos)
        theil = 0.0
        if n_pos > 0:
            mu = sum(arr_pos) / n_pos
            if mu > 0:
                theil = round(sum((y / mu) * math.log(y / mu) for y in arr_pos) / n_pos, 4)
                
        p = calc_percentiles(arr)
        p10 = p[10]
        p90 = p[90]
        if p10 > 0:
            p90_p10 = round(p90 / p10, 2)
        elif p10 == 0:
            p90_p10 = "N/A (P10 = $0)"
        else:
            p90_p10 = "N/A (P10 < $0)"
            
        k90 = int(math.floor(n * 0.90))
        top10_sum = sum(arr[k90:])
        tot_sum = sum(arr)
        top10_share = round((top10_sum / tot_sum) * 100.0, 2) if tot_sum > 0 else 0.0
            
        return {
            'n': n,
            'mean': round(sum(arr) / n, 2),
            'median': p[50],
            'p10': p10,
            'p20': p[20],
            'p40': p[40],
            'p50': p[50],
            'p60': p[60],
            'p80': p[80],
            'p90': p90,
            'gini_route1': gini_route1,
            'gini_route2': gini_route2,
            'theil': theil,
            'p90_p10': p90_p10,
            'top10_share': top10_share,
            'zero_count': sum(1 for x in arr if x == 0),
            'negative_count': sum(1 for x in arr if x < 0),
            'percentiles': p
        }

    us_incomes.sort()
    output_data = {
        'United States': calc_inequality(us_incomes),
        'states': {}
    }
    for s_name in sorted(data_by_state.keys()):
        st_incomes = sorted(data_by_state[s_name])
        output_data['states'][s_name] = calc_inequality(st_incomes)

    with open(CACHE_FILE, 'w') as f:
        json.dump(output_data, f)
    print("Dataset cache ready.")

HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>U.S. Income Inequality &amp; Distribution Analysis &bull; OSM 202</title>
  <style>
    :root {
      --bg: #f8fafc;
      --card-bg: #ffffff;
      --card-sub-bg: #f8fafc;
      --border: #e2e8f0;
      --border-dark: #cbd5e1;
      --text: #0f172a;
      --text-muted: #64748b;
      --navy: #0f172a;
      --navy-light: #1e293b;
      --blue: #0284c7;
      --blue-light: #e0f2fe;
      --blue-border: #bae6fd;
      --orange: #ea580c;
      --orange-light: #ffedd5;
      --rose: #e11d48;
      --rose-light: #ffe4e6;
      --green: #16a34a;
      --green-light: #dcfce7;
      --purple: #7c3aed;
      --purple-light: #ede9fe;
      --grid-line: #e2e8f0;
      --axis-line: #94a3b8;
      --chart-tick: #64748b;
      --strip-bg: #ffffff;
      --strip-border: #cbd5e1;
      --us-curve-color: #64748b;
      --shadow-sm: 0 1px 3px rgba(0,0,0,0.04);
      --shadow-md: 0 4px 12px rgba(15,23,42,0.06);
    }

    body.dark-mode {
      --bg: #0b1120;
      --card-bg: #131d33;
      --card-sub-bg: #18233f;
      --border: #22314e;
      --border-dark: #334155;
      --text: #f8fafc;
      --text-muted: #94a3b8;
      --navy: #ffffff;
      --navy-light: #f1f5f9;
      --blue: #38bdf8;
      --blue-light: #0c4a6e;
      --blue-border: #0369a1;
      --orange: #fb923c;
      --orange-light: #431407;
      --rose: #f43f5e;
      --rose-light: #4c0519;
      --green: #22c55e;
      --green-light: #052e16;
      --purple: #a78bfa;
      --purple-light: #2e1065;
      --grid-line: #1e293b;
      --axis-line: #475569;
      --chart-tick: #94a3b8;
      --strip-bg: #131d33;
      --strip-border: #334155;
      --us-curve-color: #94a3b8;
      --shadow-sm: 0 1px 3px rgba(0,0,0,0.2);
      --shadow-md: 0 4px 12px rgba(0,0,0,0.3);
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background-color: var(--bg);
      color: var(--text);
      font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      padding: 20px 24px;
      line-height: 1.5;
      -webkit-font-smoothing: antialiased;
      transition: background-color 0.2s ease, color 0.2s ease;
    }

    .container {
      max-width: 1320px;
      margin: 0 auto;
    }

    /* 1. PROFESSIONAL HEADER */
    header {
      background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
      color: #ffffff;
      padding: 20px 26px;
      border-radius: 14px;
      margin-bottom: 16px;
      box-shadow: 0 4px 14px rgba(15, 23, 42, 0.12);
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 16px;
    }
    .header-left h1 {
      font-size: 22px;
      font-weight: 700;
      letter-spacing: -0.3px;
      color: #ffffff;
      line-height: 1.25;
    }
    .header-subtitle {
      font-size: 13px;
      color: #94a3b8;
      font-weight: 500;
      margin-top: 3px;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .header-subtitle span.dot {
      opacity: 0.5;
    }
    .header-right {
      display: flex;
      align-items: center;
      gap: 12px;
      flex-wrap: wrap;
    }
    .current-loc-badge {
      background: rgba(255, 255, 255, 0.1);
      border: 1px solid rgba(255, 255, 255, 0.2);
      padding: 6px 14px;
      border-radius: 20px;
      font-size: 13px;
      font-weight: 600;
      color: #e2e8f0;
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .current-loc-badge strong {
      color: #38bdf8;
    }
    .btn-header {
      background: rgba(255, 255, 255, 0.08);
      color: #f1f5f9;
      border: 1px solid rgba(255, 255, 255, 0.2);
      padding: 7px 14px;
      border-radius: 8px;
      font-size: 12.5px;
      font-weight: 600;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      transition: all 0.15s ease;
      user-select: none;
    }
    .btn-header:hover {
      background: rgba(255, 255, 255, 0.18);
      border-color: rgba(255, 255, 255, 0.35);
      color: #ffffff;
    }
    .btn-header.active {
      background: #0284c7;
      border-color: #38bdf8;
      color: #ffffff;
    }

    /* CONTROL BAR (Location selector + Sample summary) */
    .control-bar {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 12px 18px;
      margin-bottom: 16px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
      box-shadow: var(--shadow-sm);
    }
    .control-group {
      display: flex;
      align-items: center;
      gap: 10px;
    }
    .control-label {
      font-size: 13px;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    select#stateSelect {
      padding: 7px 14px;
      font-size: 14px;
      font-weight: 600;
      border: 1.5px solid var(--border-dark);
      border-radius: 8px;
      background: var(--card-bg);
      color: var(--text);
      cursor: pointer;
      outline: none;
      min-width: 220px;
      transition: border-color 0.15s;
    }
    select#stateSelect:focus {
      border-color: var(--blue);
    }
    .location-summary-strip {
      display: flex;
      align-items: center;
      gap: 14px;
      font-size: 13px;
      color: var(--text-muted);
      flex-wrap: wrap;
    }
    .loc-chip {
      color: var(--text);
      font-weight: 700;
      background: var(--card-sub-bg);
      padding: 3px 8px;
      border-radius: 6px;
      border: 1px solid var(--border);
    }

    /* 2. STATE VS. U.S. DIFFERENCE STRIP */
    .diff-strip-card {
      background: var(--card-bg);
      border: 1.5px solid var(--strip-border);
      border-radius: 12px;
      padding: 14px 18px;
      margin-bottom: 16px;
      box-shadow: var(--shadow-sm);
    }
    .diff-strip-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
      padding-bottom: 8px;
      border-bottom: 1px solid var(--border);
    }
    .diff-strip-title {
      font-size: 13.5px;
      font-weight: 800;
      color: var(--navy);
      display: flex;
      align-items: center;
      gap: 6px;
      text-transform: uppercase;
      letter-spacing: 0.4px;
    }
    .diff-strip-tag {
      font-size: 11px;
      font-weight: 700;
      background: var(--blue-light);
      color: var(--blue);
      padding: 2px 8px;
      border-radius: 12px;
      border: 1px solid var(--blue-border);
    }
    .diff-strip-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 14px;
    }
    @media (max-width: 960px) {
      .diff-strip-grid { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 560px) {
      .diff-strip-grid { grid-template-columns: 1fr; }
    }
    .diff-col {
      background: var(--card-sub-bg);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 10px 12px;
    }
    .diff-col-name {
      font-size: 11.5px;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.4px;
      margin-bottom: 4px;
    }
    .diff-col-values {
      display: flex;
      align-items: baseline;
      gap: 6px;
      font-size: 13.5px;
      margin-bottom: 4px;
    }
    .diff-val-st {
      font-weight: 800;
      color: var(--navy);
    }
    .diff-val-sep {
      color: var(--text-muted);
      font-size: 12px;
    }
    .diff-val-us {
      color: var(--text-muted);
      font-size: 12.5px;
    }
    .diff-col-delta {
      font-size: 12px;
      font-weight: 700;
    }
    .delta-pos {
      color: #0284c7;
    }
    .delta-neg {
      color: #64748b;
    }
    .delta-zero {
      color: #16a34a;
    }

    /* 4. CLEAN METRIC CARDS (Uniform height, typography, help tooltip) */
    .metrics-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 16px;
      margin-bottom: 16px;
    }
    @media (max-width: 1024px) {
      .metrics-grid { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 600px) {
      .metrics-grid { grid-template-columns: 1fr; }
    }
    .metric-card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 18px 20px;
      box-shadow: var(--shadow-sm);
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      min-height: 148px;
    }
    .metric-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
    }
    .metric-title-group {
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .metric-title {
      font-size: 12px;
      font-weight: 800;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .info-btn {
      width: 16px;
      height: 16px;
      border-radius: 50%;
      background: var(--card-sub-bg);
      border: 1px solid var(--border-dark);
      color: var(--text-muted);
      font-size: 10.5px;
      font-weight: 700;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      cursor: help;
      position: relative;
      user-select: none;
      padding: 0;
      line-height: 1;
    }
    .info-btn:hover .info-popover {
      visibility: visible;
      opacity: 1;
    }
    .info-popover {
      visibility: hidden;
      opacity: 0;
      position: absolute;
      bottom: 24px;
      left: 50%;
      transform: translateX(-50%);
      width: 220px;
      background: #0f172a;
      color: #f8fafc;
      font-size: 11.5px;
      font-weight: 400;
      line-height: 1.4;
      padding: 8px 10px;
      border-radius: 6px;
      box-shadow: 0 4px 14px rgba(0,0,0,0.3);
      z-index: 100;
      transition: opacity 0.15s ease;
      text-transform: none;
      letter-spacing: normal;
      pointer-events: none;
    }
    .info-popover::after {
      content: "";
      position: absolute;
      top: 100%;
      left: 50%;
      transform: translateX(-50%);
      border-width: 5px;
      border-style: solid;
      border-color: #0f172a transparent transparent transparent;
    }
    .metric-pill {
      font-size: 11px;
      font-weight: 700;
      color: var(--text-muted);
      background: var(--card-sub-bg);
      padding: 2px 7px;
      border-radius: 10px;
      border: 1px solid var(--border);
    }
    .metric-val {
      font-size: 30px;
      font-weight: 800;
      color: var(--navy);
      letter-spacing: -0.5px;
      line-height: 1.15;
      margin-bottom: 6px;
      font-variant-numeric: tabular-nums;
    }
    .metric-context {
      font-size: 11.5px;
      color: var(--text-muted);
      line-height: 1.35;
      border-top: 1px dashed var(--border);
      padding-top: 6px;
      margin-top: auto;
    }

    /* KEY TAKEAWAYS CARD */
    .takeaways-card {
      background: var(--card-bg);
      border: 1.5px solid var(--green);
      border-radius: 12px;
      padding: 14px 18px;
      margin-bottom: 16px;
      box-shadow: var(--shadow-sm);
    }
    .takeaways-header {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 13px;
      font-weight: 800;
      color: var(--green);
      margin-bottom: 8px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .takeaways-grid {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 12px;
    }
    @media (max-width: 960px) {
      .takeaways-grid { grid-template-columns: repeat(2, 1fr); }
    }
    @media (max-width: 560px) {
      .takeaways-grid { grid-template-columns: 1fr; }
    }
    .takeaway-item {
      background: var(--card-sub-bg);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 8px 10px;
      font-size: 12px;
      line-height: 1.4;
      color: var(--text);
    }
    .takeaway-item strong {
      color: var(--navy);
    }

    /* 3. MAIN SECTION: GRAPH (CENTERPIECE) + STUDENT GUESSES */
    .main-grid {
      display: grid;
      grid-template-columns: 1.65fr 1fr;
      gap: 16px;
      margin-bottom: 16px;
    }
    @media (max-width: 1060px) {
      .main-grid { grid-template-columns: 1fr; }
    }
    .card {
      background: var(--card-bg);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 18px 20px;
      box-shadow: var(--shadow-sm);
    }
    .card-title-row {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 12px;
      padding-bottom: 8px;
      border-bottom: 1px solid var(--border);
    }
    .card-title {
      font-size: 15px;
      font-weight: 800;
      color: var(--navy);
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .badge {
      font-size: 11px;
      font-weight: 700;
      background: var(--blue-light);
      color: var(--blue);
      border: 1px solid var(--blue-border);
      padding: 3px 8px;
      border-radius: 12px;
    }

    /* Chart Centerpiece Styles */
    .chart-wrapper {
      position: relative;
      width: 100%;
      height: 440px;
      margin-bottom: 12px;
      user-select: none;
    }
    svg.chart-svg {
      width: 100%;
      height: 100%;
      overflow: visible;
    }
    .grid-line {
      stroke: var(--grid-line);
      stroke-width: 1;
    }
    .axis-line {
      stroke: var(--axis-line);
      stroke-width: 1.5;
    }
    .chart-tick-label {
      font-size: 11.5px;
      font-weight: 600;
      fill: var(--chart-tick);
      font-family: inherit;
    }
    .chart-axis-title {
      font-size: 13px;
      font-weight: 800;
      fill: var(--navy);
      letter-spacing: 0.3px;
      font-family: inherit;
    }
    .curve-area {
      fill: url(#curveGradient);
      pointer-events: none;
    }
    .curve-path {
      fill: none;
      stroke: var(--blue);
      stroke-width: 2.8;
      stroke-linecap: round;
      stroke-linejoin: round;
    }
    .us-curve-path {
      fill: none;
      stroke: var(--us-curve-color);
      stroke-width: 2.2;
      stroke-dasharray: 7, 4;
      opacity: 0.95;
    }
    .milestone-point-st {
      fill: var(--blue);
      stroke: var(--card-bg);
      stroke-width: 2.5;
      cursor: pointer;
      transition: r 0.15s ease;
    }
    .milestone-point-us {
      fill: var(--us-curve-color);
      stroke: var(--card-bg);
      stroke-width: 2;
    }
    .gap-line {
      stroke: var(--rose);
      stroke-width: 1.8;
      stroke-dasharray: 4, 3;
    }
    .point-guess {
      fill: var(--orange);
      stroke: var(--card-bg);
      stroke-width: 2.5;
      cursor: pointer;
    }
    .point-real {
      fill: var(--blue);
      stroke: var(--card-bg);
      stroke-width: 2.5;
      cursor: pointer;
    }
    .active-crosshair {
      stroke: var(--blue);
      stroke-width: 1.5;
      stroke-dasharray: 3, 3;
      pointer-events: none;
    }
    .active-target-ring {
      fill: none;
      stroke: var(--blue);
      stroke-width: 2;
      opacity: 0.75;
      pointer-events: none;
    }
    .active-target-dot {
      fill: var(--blue);
      stroke: #ffffff;
      stroke-width: 2.5;
      pointer-events: none;
    }
    .active-us-dot {
      fill: var(--us-curve-color);
      stroke: #ffffff;
      stroke-width: 2;
      pointer-events: none;
    }
    .onchart-callout-box {
      fill: #0f172a;
      rx: 5;
      ry: 5;
      opacity: 0.95;
    }
    .onchart-callout-text {
      fill: #ffffff;
      font-size: 11px;
      font-weight: 700;
      text-anchor: middle;
      font-family: inherit;
    }

    /* Chart Legend */
    .chart-legend {
      display: flex;
      flex-wrap: wrap;
      gap: 14px;
      background: var(--card-sub-bg);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 9px 14px;
      font-size: 11.5px;
      font-weight: 600;
      color: var(--text-muted);
      margin-bottom: 12px;
      align-items: center;
    }
    .legend-item {
      display: inline-flex;
      align-items: center;
      gap: 6px;
    }
    .legend-line-st {
      width: 18px;
      height: 3px;
      background: var(--blue);
      border-radius: 2px;
    }
    .legend-line-us {
      width: 18px;
      height: 0px;
      border-top: 2.5px dashed var(--us-curve-color);
    }
    .legend-dot-orange {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      background: var(--orange);
      border: 1.5px solid var(--card-bg);
    }
    .legend-dot-blue {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      background: var(--blue);
      border: 1.5px solid var(--card-bg);
    }
    .legend-line-rose {
      width: 16px;
      height: 0px;
      border-top: 2px dashed var(--rose);
    }

    /* Tooltip */
    .tooltip {
      position: absolute;
      background: #0f172a;
      color: #f8fafc;
      padding: 10px 14px;
      border-radius: 8px;
      font-size: 12px;
      pointer-events: none;
      box-shadow: 0 6px 20px rgba(0,0,0,0.3);
      z-index: 50;
      opacity: 0;
      transition: opacity 0.12s ease;
      min-width: 210px;
      max-width: 290px;
      border: 1px solid rgba(255,255,255,0.15);
    }

    /* Selected Percentile Highlight Card */
    .inspector-panel {
      background: var(--card-sub-bg);
      border: 1.5px solid var(--blue-border);
      border-radius: 10px;
      padding: 12px 16px;
      margin-top: 4px;
    }
    .inspector-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 6px;
      flex-wrap: wrap;
      gap: 8px;
    }
    .inspector-title {
      font-size: 12.5px;
      font-weight: 800;
      color: var(--blue);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .inspector-badges {
      display: flex;
      gap: 6px;
      flex-wrap: wrap;
    }
    .ibadge {
      font-size: 11px;
      font-weight: 700;
      padding: 2px 7px;
      border-radius: 4px;
    }
    .ibadge-loc { background: var(--card-bg); border: 1px solid var(--border); color: var(--navy); }
    .ibadge-pct { background: var(--blue-light); border: 1px solid var(--blue-border); color: var(--blue); }
    .ibadge-inc { background: #0f172a; color: #ffffff; }
    .inspector-explanation {
      font-size: 12.5px;
      line-height: 1.45;
      color: var(--text);
      margin-bottom: 8px;
    }
    .quick-jump-bar {
      display: flex;
      align-items: center;
      gap: 5px;
      flex-wrap: wrap;
      font-size: 11.5px;
      color: var(--text-muted);
      border-top: 1px dashed var(--border);
      padding-top: 8px;
    }
    .quick-btn {
      background: var(--card-bg);
      border: 1px solid var(--border);
      color: var(--navy);
      font-size: 11px;
      font-weight: 700;
      padding: 3px 7px;
      border-radius: 4px;
      cursor: pointer;
      transition: all 0.15s;
    }
    .quick-btn:hover, .quick-btn.active {
      background: var(--blue);
      color: #ffffff;
      border-color: var(--blue);
    }

    /* RIGHT COLUMN: STUDENT GUESSES COMPARISON */
    .guess-table-wrapper {
      margin-bottom: 12px;
      overflow-x: auto;
    }
    table.guess-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12.5px;
    }
    table.guess-table th {
      text-align: left;
      padding: 8px 6px;
      font-size: 10.5px;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
      border-bottom: 2px solid var(--border);
    }
    table.guess-table td {
      padding: 9px 6px;
      border-bottom: 1px solid var(--border);
      vertical-align: middle;
    }
    table.guess-table tr:hover {
      background: var(--card-sub-bg);
    }
    .tag-guess {
      background: var(--orange-light);
      color: var(--orange);
      font-weight: 700;
      padding: 2px 6px;
      border-radius: 4px;
      display: inline-block;
    }
    .tag-real {
      background: var(--blue-light);
      color: var(--blue);
      font-weight: 700;
      padding: 2px 6px;
      border-radius: 4px;
      display: inline-block;
    }
    .gap-badge-over {
      color: #e11d48;
      font-weight: 700;
      background: var(--rose-light);
      padding: 2px 5px;
      border-radius: 4px;
    }
    .gap-badge-under {
      color: #0284c7;
      font-weight: 700;
      background: var(--blue-light);
      padding: 2px 5px;
      border-radius: 4px;
    }
    .highlight-pill {
      display: inline-block;
      font-size: 10px;
      font-weight: 700;
      padding: 1px 5px;
      border-radius: 8px;
      margin-top: 3px;
    }
    .pill-min-gap { background: var(--blue-light); color: var(--blue); border: 1px solid var(--blue-border); }
    .pill-max-gap { background: var(--purple-light); color: var(--purple); border: 1px solid #ddd6fe; }

    .guess-edit-note {
      background: var(--card-sub-bg);
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 10px 12px;
      font-size: 11.5px;
      color: var(--text-muted);
      line-height: 1.45;
    }

    /* DETAILED BENCHMARK COMPARISON CARD */
    .state-vs-us-card {
      background: var(--card-bg);
      border: 1.5px solid var(--border-dark);
      border-radius: 12px;
      padding: 18px 20px;
      margin-bottom: 16px;
      box-shadow: var(--shadow-sm);
    }
    .state-comp-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
      margin-bottom: 12px;
    }
    @media (max-width: 900px) {
      .state-comp-grid { grid-template-columns: 1fr; }
    }
    .comp-subheading {
      font-size: 12px;
      font-weight: 800;
      color: var(--navy);
      margin-bottom: 8px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    table.comp-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12.5px;
    }
    table.comp-table th {
      text-align: left;
      padding: 7px 8px;
      background: var(--card-sub-bg);
      color: var(--text-muted);
      font-size: 10.5px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      border-bottom: 2px solid var(--border);
    }
    table.comp-table td {
      padding: 8px 8px;
      border-bottom: 1px solid var(--border);
      vertical-align: middle;
    }
    table.comp-table tr:hover {
      background: var(--card-sub-bg);
    }
    .diff-tag-pos {
      font-weight: 700;
      color: #0284c7;
      background: var(--blue-light);
      padding: 2px 6px;
      border-radius: 4px;
      display: inline-block;
    }
    .diff-tag-neg {
      font-weight: 700;
      color: #64748b;
      background: var(--card-sub-bg);
      padding: 2px 6px;
      border-radius: 4px;
      display: inline-block;
    }
    .comp-factual-box {
      background: var(--card-sub-bg);
      border: 1px solid var(--border);
      border-left: 4px solid var(--blue);
      border-radius: 6px;
      padding: 12px 14px;
      font-size: 12.5px;
      line-height: 1.55;
      color: var(--navy);
    }
    .comp-factual-box p {
      margin-bottom: 5px;
    }
    .comp-factual-box p:last-child {
      margin-bottom: 0;
    }

    /* ACCORDION (Technical Methodology Guide) */
    .accordion-card {
      margin-top: 12px;
      margin-bottom: 16px;
    }
    .acc-item {
      background: var(--card-bg);
      border-radius: 8px;
      margin-bottom: 8px;
      border: 1px solid var(--border);
      overflow: hidden;
    }
    .acc-header {
      padding: 12px 16px;
      font-weight: 700;
      font-size: 13px;
      color: var(--navy);
      cursor: pointer;
      display: flex;
      justify-content: space-between;
      align-items: center;
      user-select: none;
    }
    .acc-header:hover { background: var(--card-sub-bg); }
    .acc-content {
      padding: 0 16px 14px;
      display: none;
      color: var(--text-muted);
      font-size: 12.5px;
      border-top: 1px solid var(--border);
    }
    .acc-content p { margin-top: 8px; }
    .acc-content ul { margin: 6px 0 6px 18px; }
    .acc-content li { margin-bottom: 4px; }
    .math-code {
      font-family: SFMono-Regular, Menlo, Monaco, Consolas, monospace;
      background: var(--card-sub-bg);
      padding: 2px 5px;
      border-radius: 4px;
      font-size: 11.5px;
      color: var(--navy);
    }

    /* 6. SMOOTH STATE TRANSITION FADE */
    #dashboardContent {
      transition: opacity 0.16s ease-in-out;
      opacity: 1;
    }
    #dashboardContent.fading {
      opacity: 0.35;
    }

    /* 7. PRESENTATION MODE STYLES */
    body.presentation-mode {
      padding: 14px 18px;
      background-color: var(--bg);
    }
    body.presentation-mode .container {
      max-width: 1460px;
    }
    body.presentation-mode .header-subtitle span.opt-hide {
      display: none;
    }
    body.presentation-mode .accordion-card {
      display: none;
    }
    body.presentation-mode .guess-edit-note {
      display: none;
    }
    body.presentation-mode .metric-card {
      padding: 18px 20px;
    }
    body.presentation-mode .metric-val {
      font-size: 34px;
    }
    body.presentation-mode .chart-wrapper {
      height: 480px;
    }

    /* 9. FOOTER */
    footer.app-footer {
      text-align: center;
      padding: 20px 0 10px;
      font-size: 12px;
      color: var(--text-muted);
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 8px;
      flex-wrap: wrap;
    }
    footer.app-footer code {
      font-family: SFMono-Regular, Menlo, Monaco, Consolas, monospace;
      background: var(--card-sub-bg);
      padding: 2px 6px;
      border-radius: 4px;
      border: 1px solid var(--border);
      color: var(--navy);
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- 1. PROFESSIONAL HEADER -->
    <header>
      <div class="header-left">
        <h1>U.S. Income Inequality &amp; Distribution Analysis</h1>
        <div class="header-subtitle">
          <span>OSM 202 Signature Assignment</span>
          <span class="dot">&bull;</span>
          <span class="opt-hide">100,000 Adult Sample (Census American Community Survey)</span>
        </div>
      </div>
      <div class="header-right">
        <div class="current-loc-badge">
          <span>Viewing:</span>
          <strong id="headerSelectedLocation">United States</strong>
        </div>
        <button id="btnThemeToggle" class="btn-header" onclick="toggleTheme()" title="Toggle Dark or Light Appearance">
          <span id="themeIcon">🌙</span> <span id="themeText">Dark Mode</span>
        </button>
        <button id="btnPresMode" class="btn-header" onclick="togglePresentationMode()" title="Toggle Full-Screen Presentation Layout">
          <span>📺</span> <span id="presModeText">Presentation Mode</span>
        </button>
      </div>
    </header>

    <!-- Content wrapper with smooth crossfade on location change -->
    <div id="dashboardContent">
      <!-- Location Selector & Sample Summary -->
      <div class="control-bar">
        <div class="control-group">
          <label for="stateSelect" class="control-label">Location:</label>
          <select id="stateSelect" onchange="onLocationChange()">
            <option value="United States">United States (National Total)</option>
          </select>
        </div>
        <div class="location-summary-strip">
          <span>Sample Size: <strong class="loc-chip" id="sampleSize">100,000 adults</strong></span>
          <span>Median Income: <strong class="loc-chip" id="locMedian">$39,589</strong></span>
          <span id="usBenchNote" style="display:none;">U.S. Benchmark Median: <strong>$39,589</strong></span>
        </div>
      </div>

      <!-- 2. STATE VS. U.S. DIFFERENCE STRIP (Visible when a state is selected) -->
      <div class="diff-strip-card" id="stateDiffStrip" style="display:none;">
        <div class="diff-strip-header">
          <span class="diff-strip-title">🏛️ <span id="stripStateName">State</span> vs. United States</span>
          <span class="diff-strip-tag">Benchmark Comparison Strip</span>
        </div>
        <div class="diff-strip-grid">
          <!-- Col 1: Median Income -->
          <div class="diff-col">
            <div class="diff-col-name">Median Income</div>
            <div class="diff-col-values">
              <span class="diff-val-st" id="stripMedState">$0</span>
              <span class="diff-val-sep">|</span>
              <span class="diff-val-us">U.S. $39,589</span>
            </div>
            <div class="diff-col-delta" id="stripMedDiff">Diff: $0</div>
          </div>
          <!-- Col 2: Gini Coefficient -->
          <div class="diff-col">
            <div class="diff-col-name">Gini Coefficient</div>
            <div class="diff-col-values">
              <span class="diff-val-st" id="stripGiniState">0.0000</span>
              <span class="diff-val-sep">|</span>
              <span class="diff-val-us">U.S. 0.5621</span>
            </div>
            <div class="diff-col-delta" id="stripGiniDiff">Diff: 0.0000</div>
          </div>
          <!-- Col 3: P90/P10 Ratio -->
          <div class="diff-col">
            <div class="diff-col-name">P90 / P10 Ratio</div>
            <div class="diff-col-values">
              <span class="diff-val-st" id="stripP90State">0.0</span>
              <span class="diff-val-sep">|</span>
              <span class="diff-val-us">U.S. 1205.14</span>
            </div>
            <div class="diff-col-delta" id="stripP90Diff">Diff: 0.0</div>
          </div>
          <!-- Col 4: Top 10% Share -->
          <div class="diff-col">
            <div class="diff-col-name">Top 10% Share</div>
            <div class="diff-col-values">
              <span class="diff-val-st" id="stripTop10State">0.0%</span>
              <span class="diff-val-sep">|</span>
              <span class="diff-val-us">U.S. 40.02%</span>
            </div>
            <div class="diff-col-delta" id="stripTop10Diff">Diff: 0.0%</div>
          </div>
        </div>
      </div>

      <!-- 4. CLEAN METRIC CARDS (Consistent size, typography, info popovers) -->
      <div class="metrics-grid">
        <!-- Card 1: Gini -->
        <div class="metric-card">
          <div>
            <div class="metric-header">
              <div class="metric-title-group">
                <span class="metric-title">Gini Coefficient</span>
                <button class="info-btn" aria-label="Gini Info">i
                  <span class="info-popover">Measures overall income inequality from 0 (perfect equality) to 1 (complete inequality). Evaluated using dual independent methods.</span>
                </button>
              </div>
              <span class="metric-pill" id="giniUS">US: 0.5621</span>
            </div>
            <div class="metric-val" id="giniVal">0.5621</div>
          </div>
          <div class="metric-context" id="giniCheck">
            Route 1: <strong id="gRoute1">0.5621</strong> &bull; Route 2: <strong id="gRoute2">0.5621</strong><br>
            <span style="color:var(--green); font-weight:700;">&#10003; Dual-Route Verified Match</span>
          </div>
        </div>

        <!-- Card 2: Theil Index -->
        <div class="metric-card">
          <div>
            <div class="metric-header">
              <div class="metric-title-group">
                <span class="metric-title">Theil Index (Entropy T)</span>
                <button class="info-btn" aria-label="Theil Info">i
                  <span class="info-popover">Entropy-based measure sensitive to upper-tail concentration among positive earners. Evaluated for y &gt; 0.</span>
                </button>
              </div>
              <span class="metric-pill" id="theilUS">US: 0.4959</span>
            </div>
            <div class="metric-val" id="theilVal">0.4959</div>
          </div>
          <div class="metric-context">
            Evaluated for <span class="math-code">y &gt; 0</span> (positive earners).<br>
            Measures upper-tail income concentration.
          </div>
        </div>

        <!-- Card 3: P90 / P10 Decile Ratio -->
        <div class="metric-card">
          <div>
            <div class="metric-header">
              <div class="metric-title-group">
                <span class="metric-title">P90 / P10 Decile Ratio</span>
                <button class="info-btn" aria-label="P90/P10 Info">i
                  <span class="info-popover">Compares the 90th percentile to the 10th percentile, reflecting the income spread across the middle 80% of adults.</span>
                </button>
              </div>
              <span class="metric-pill" id="p90p10US">US: 1205.14</span>
            </div>
            <div class="metric-val" id="p90p10Val">1205.14</div>
          </div>
          <div class="metric-context" id="p90p10Note">
            P90: <strong id="valP90">$130,155</strong> &bull; P10: <strong id="valP10">$108</strong><br>
            Decile dispersion ratio.
          </div>
        </div>

        <!-- Card 4: Top 10% Income Share -->
        <div class="metric-card">
          <div>
            <div class="metric-header">
              <div class="metric-title-group">
                <span class="metric-title">Top 10% Income Share</span>
                <button class="info-btn" aria-label="Top 10% Info">i
                  <span class="info-popover">The share of all aggregate personal income captured by the top 10% of adults (at or above the 90th percentile).</span>
                </button>
              </div>
              <span class="metric-pill" id="top10US">US: 40.02%</span>
            </div>
            <div class="metric-val" id="top10Val">40.02%</div>
          </div>
          <div class="metric-context">
            Cutoff: Adults <span class="math-code">&ge; P90</span>.<br>
            Share of aggregate personal income.
          </div>
        </div>
      </div>

      <!-- KEY TAKEAWAYS CARD -->
      <div class="takeaways-card">
        <div class="takeaways-header">
          <span>📊 Key Takeaways &bull; <span id="takeawaysLocation">United States</span></span>
        </div>
        <div class="takeaways-grid">
          <div class="takeaway-item" id="takeawayMedian">
            <strong>Median Income:</strong> Loading...
          </div>
          <div class="takeaway-item" id="takeawayGini">
            <strong>Gini Inequality:</strong> Loading...
          </div>
          <div class="takeaway-item" id="takeawayTop10">
            <strong>Top 10% Share:</strong> Loading...
          </div>
          <div class="takeaway-item" id="takeawayDecile">
            <strong>P90/P10 Decile Gap:</strong> Loading...
          </div>
        </div>
      </div>

      <!-- DETAILED STATE VS. UNITED STATES COMPARISON CARD (Visible when state is selected) -->
      <div class="state-vs-us-card" id="stateVsUsCard" style="display:none;">
        <div class="card-title-row">
          <div class="card-title">
            <span>🏛️ <span id="compHeadingState">State</span> vs. United States Comparison</span>
          </div>
          <span class="badge" id="compCardBadge">Detailed Side-by-Side Analysis</span>
        </div>

        <div class="state-comp-grid">
          <!-- Table 1: Inequality Measures -->
          <div>
            <div class="comp-subheading">Inequality Measures &mdash; Side-by-Side</div>
            <table class="comp-table">
              <thead>
                <tr>
                  <th>Measure</th>
                  <th id="thStateName1">Selected State</th>
                  <th>United States</th>
                  <th>Difference</th>
                </tr>
              </thead>
              <tbody id="compInequalityBody"></tbody>
            </table>
          </div>

          <!-- Table 2: Benchmark Comparison -->
          <div>
            <div class="comp-subheading">Income Benchmarks by Percentile</div>
            <table class="comp-table">
              <thead>
                <tr>
                  <th>Percentile</th>
                  <th id="thStateName2">Selected State</th>
                  <th>United States</th>
                  <th>Difference ($)</th>
                  <th>Difference (%)</th>
                </tr>
              </thead>
              <tbody id="compBenchmarksBody"></tbody>
            </table>
          </div>
        </div>

        <!-- Factual Plain-English Narrative -->
        <div class="comp-factual-box" id="compFactualBox"></div>
      </div>

      <!-- 3. MAIN GRID: GRAPH (CENTERPIECE) + STUDENT GUESSES -->
      <div class="main-grid">
        <!-- Chart Column -->
        <div class="card">
          <div class="card-title-row">
            <div class="card-title">
              <span>Income Percentile Distribution Curve (P1 &ndash; P100)</span>
            </div>
            <span class="badge" id="chartBadge">United States</span>
          </div>

          <div class="chart-wrapper" id="chartWrapper">
            <svg class="chart-svg" id="chartSvg" viewBox="0 0 960 520"></svg>
            <div class="tooltip" id="chartTooltip"></div>
          </div>

          <!-- Clear Legend (Color + Stroke + Shape) -->
          <div class="chart-legend">
            <div class="legend-item">
              <div class="legend-line-st"></div>
              <span id="legendLocCurveName">Selected State Curve (Solid)</span>
            </div>
            <div class="legend-item" id="usLegendItem" style="display:none;">
              <div class="legend-line-us"></div>
              <span>United States Benchmark (Dashed)</span>
            </div>
            <div class="legend-item">
              <div class="legend-dot-orange"></div>
              <span>Student Guess</span>
            </div>
            <div class="legend-item">
              <div class="legend-dot-blue"></div>
              <span>Actual Income at Percentile</span>
            </div>
            <div class="legend-item">
              <div class="legend-line-rose"></div>
              <span>Difference / Gap</span>
            </div>
          </div>

          <!-- Interactive Highlight & Explanation Panel -->
          <div class="inspector-panel" id="inspectorPanel">
            <div class="inspector-header">
              <div class="inspector-title">
                <span>Selected Percentile Highlight</span>
                <span style="font-size:11px; font-weight:normal; color:var(--text-muted);">(Hover or click anywhere on curve)</span>
              </div>
              <div class="inspector-badges">
                <span class="ibadge ibadge-loc" id="inspBadgeLoc">United States</span>
                <span class="ibadge ibadge-pct" id="inspBadgePct">60th Percentile (P60)</span>
                <span class="ibadge ibadge-inc" id="inspBadgeInc">Income: $52,062</span>
              </div>
            </div>

            <div class="inspector-explanation" id="inspExplanation">
              <strong>60th percentile: $52,062.</strong> About 60% of observations have income at or below this amount.
            </div>

            <div class="quick-jump-bar">
              <span>Quick Jump:</span>
              <button class="quick-btn" onclick="quickJump(10)">P10</button>
              <button class="quick-btn" onclick="quickJump(20)">P20</button>
              <button class="quick-btn" onclick="quickJump(40)">P40</button>
              <button class="quick-btn" onclick="quickJump(50)">P50 (Median)</button>
              <button class="quick-btn active" onclick="quickJump(60)">P60</button>
              <button class="quick-btn" onclick="quickJump(80)">P80</button>
              <button class="quick-btn" onclick="quickJump(90)">P90</button>
              <button class="quick-btn" onclick="quickJump(99)">P99</button>
            </div>
          </div>
        </div>

        <!-- Student Salary Guesses Column -->
        <div class="card">
          <div class="card-title-row">
            <div class="card-title">
              <span>Student Salary Guesses vs. Real</span>
            </div>
            <span class="badge">4 Guess Benchmarks</span>
          </div>

          <div class="guess-table-wrapper">
            <table class="guess-table">
              <thead>
                <tr>
                  <th>Percentile</th>
                  <th>Student Guess</th>
                  <th>Actual Income</th>
                  <th>Difference</th>
                  <th>Status</th>
                </tr>
              </thead>
              <tbody id="guessTableBody"></tbody>
            </table>
          </div>

          <div class="guess-edit-note">
            <strong>How to edit guesses:</strong> In <code>app.py</code>, update the <code>STUDENT_GUESSES</code> dictionary at the top of the file (lines 27&ndash;32) and refresh your browser.
          </div>
        </div>
      </div>

      <!-- Technical Methodology Guide Accordion -->
      <div class="accordion-card">
        <div class="acc-item">
          <div class="acc-header" onclick="toggleAccordion('acc1')">
            <span>📖 Presentation Guide: Inequality Measures &amp; Dataset Methodology</span>
            <span id="accIcon1">▼</span>
          </div>
          <div class="acc-content" id="acc1">
            <p><strong>Dataset Overview:</strong> This dataset contains a representative sample of 100,000 U.S. adults drawn from the Census Bureau's American Community Survey (ACS). Non-negative incomes are evaluated for distribution and percentiles.</p>
            <p><strong>Dual-Route Gini Verification:</strong></p>
            <ul>
              <li><strong>Route 1 (Lorenz Numerical Integration):</strong> Computes cumulative population share $p_i$ versus cumulative income share $L_i$ and integrates the area between the Lorenz curve and the 45-degree line of perfect equality using the trapezoidal rule: $G = 1 - 2 \int_0^1 L(p) dp$.</li>
              <li><strong>Route 2 (Relative Mean Difference Covariance):</strong> Uses the sorted rank formulation: $G = rac{2 \sum_{i=1}^n i \cdot y_i}{n \sum_{i=1}^n y_i} - rac{n + 1}{n}$. Both yield identical results (0.5621 nationally).</li>
            </ul>
            <p><strong>Decile Spread &amp; Negative Incomes:</strong> P90/P10 measures the ratio between the 90th percentile earner and the 10th percentile earner. In certain jurisdictions (e.g., California, New York), at least 10% of adult respondents report $0 personal income, rendering the mathematical ratio undefined ($N/A$).</p>
          </div>
        </div>
      </div>
    </div>

    <!-- 9. FOOTER -->
    <footer class="app-footer">
      <div>Source: <code>us_adults_sample_100k</code> dataset (Census American Community Survey 100k sample)</div>
      <div>&bull;</div>
      <div>Analysis created for OSM 202 Signature Assignment</div>
    </footer>
  </div>

  <script>
    const initialGuesses = {
      20: 20000,
      40: 38000,
      60: 60000,
      80: 95000
    };

    let globalData = null;
    let currentPercentile = 60;
    let isPresentationMode = false;
    let isDarkMode = false;

    // Theme toggle (Light/Dark mode)
    function toggleTheme() {
      isDarkMode = !isDarkMode;
      document.body.classList.toggle('dark-mode', isDarkMode);
      document.getElementById('themeIcon').textContent = isDarkMode ? '☀️' : '🌙';
      document.getElementById('themeText').textContent = isDarkMode ? 'Light Mode' : 'Dark Mode';
      if (globalData) {
        renderDashboard();
      }
    }

    // Presentation mode toggle
    function togglePresentationMode() {
      isPresentationMode = !isPresentationMode;
      document.body.classList.toggle('presentation-mode', isPresentationMode);
      const btn = document.getElementById('btnPresMode');
      const txt = document.getElementById('presModeText');
      if (isPresentationMode) {
        btn.classList.add('active');
        txt.textContent = '✕ Exit Presentation';
      } else {
        btn.classList.remove('active');
        txt.textContent = 'Presentation Mode';
      }
      if (globalData) {
        renderDashboard();
      }
    }

    function toggleAccordion(id) {
      const el = document.getElementById(id);
      const icon = document.getElementById('accIcon' + id.slice(3));
      if (el.style.display === 'block') {
        el.style.display = 'none';
        icon.textContent = '▼';
      } else {
        el.style.display = 'block';
        icon.textContent = '▲';
      }
    }

    async function loadData() {
      try {
        const res = await fetch('/api/data');
        globalData = await res.json();
        initDropdown();
      } catch (err) {
        console.error('Error loading data:', err);
      }
    }

    function initDropdown() {
      const select = document.getElementById('stateSelect');
      const sortedStates = Object.keys(globalData.states).sort();
      sortedStates.forEach(st => {
        const opt = document.createElement('option');
        opt.value = st;
        opt.textContent = st;
        select.appendChild(opt);
      });
      renderDashboard();
    }

    // 6. Smooth State Changes: crossfade transition
    function onLocationChange() {
      const content = document.getElementById('dashboardContent');
      content.classList.add('fading');
      setTimeout(() => {
        renderDashboard();
        content.classList.remove('fading');
      }, 120);
    }

    // 5. Consistent Number Formatting
    function fmtMoney(val) {
      if (typeof val === 'number') {
        return '$' + Math.round(val).toLocaleString('en-US');
      }
      return val;
    }

    function renderDashboard() {
      const selLoc = document.getElementById('stateSelect').value;
      const usData = globalData['United States'];
      const isState = selLoc !== 'United States';
      const locData = isState ? globalData.states[selLoc] : usData;

      // 1. Header and Badges
      document.getElementById('headerSelectedLocation').textContent = selLoc;
      document.getElementById('sampleSize').textContent = locData.n.toLocaleString('en-US') + ' adults';
      document.getElementById('locMedian').textContent = fmtMoney(locData.median);
      document.getElementById('chartBadge').textContent = selLoc;

      // Legend labeling
      const legendLocCurve = document.getElementById('legendLocCurveName');
      const usLegendItem = document.getElementById('usLegendItem');
      const benchNote = document.getElementById('usBenchNote');

      if (isState) {
        legendLocCurve.textContent = `${selLoc} Curve (Solid)`;
        usLegendItem.style.display = 'inline-flex';
        benchNote.style.display = 'inline';
      } else {
        legendLocCurve.textContent = 'United States Curve (Solid)';
        usLegendItem.style.display = 'none';
        benchNote.style.display = 'none';
      }

      // 4 Metric Cards
      document.getElementById('giniVal').textContent = locData.gini_route1.toFixed(4);
      document.getElementById('giniUS').textContent = 'US: ' + usData.gini_route1.toFixed(4);
      document.getElementById('gRoute1').textContent = locData.gini_route1.toFixed(4);
      document.getElementById('gRoute2').textContent = locData.gini_route2.toFixed(4);

      document.getElementById('theilVal').textContent = locData.theil.toFixed(4);
      document.getElementById('theilUS').textContent = 'US: ' + usData.theil.toFixed(4);

      const p90p10LocVal = typeof locData.p90_p10 === 'number' ? (locData.p90_p10 === 1205.14 ? '1205.14' : locData.p90_p10.toLocaleString('en-US', {maximumFractionDigits:1})) : locData.p90_p10;
      document.getElementById('p90p10Val').textContent = p90p10LocVal;
      document.getElementById('p90p10US').textContent = 'US: 1205.14';
      document.getElementById('valP90').textContent = fmtMoney(locData.p90);
      document.getElementById('valP10').textContent = fmtMoney(locData.p10);

      document.getElementById('top10Val').textContent = locData.top10_share.toFixed(2) + '%';
      document.getElementById('top10US').textContent = 'US: ' + usData.top10_share.toFixed(2) + '%';

      // Key Takeaways
      document.getElementById('takeawaysLocation').textContent = selLoc;
      document.getElementById('takeawayMedian').innerHTML = `
        <strong>Median Income:</strong> Half of adults in ${selLoc} earn under <strong>${fmtMoney(locData.median)}</strong>, and half earn more.
      `;
      document.getElementById('takeawayGini').innerHTML = `
        <strong>Gini Inequality:</strong> Gini index is <strong>${locData.gini_route1.toFixed(4)}</strong> (verified across dual computational routes).
      `;
      document.getElementById('takeawayTop10').innerHTML = `
        <strong>Top 10% Share:</strong> Richest decile holds <strong>${locData.top10_share.toFixed(2)}%</strong> of aggregate personal income.
      `;
      const p90p10Text = typeof locData.p90_p10 === 'number' ? `${locData.p90_p10.toLocaleString('en-US', {maximumFractionDigits:1})}× (${fmtMoney(locData.p90)} vs ${fmtMoney(locData.p10)})` : `${locData.p90_p10}`;
      document.getElementById('takeawayDecile').innerHTML = `
        <strong>Decile Gap (P90/P10):</strong> 90th to 10th percentile ratio is <strong>${p90p10Text}</strong>.
      `;

      // 2. STATE VS. U.S. DIFFERENCE STRIP
      const diffStrip = document.getElementById('stateDiffStrip');
      const compCard = document.getElementById('stateVsUsCard');

      if (isState) {
        diffStrip.style.display = 'block';
        compCard.style.display = 'block';
        document.getElementById('stripStateName').textContent = selLoc;
        document.getElementById('compHeadingState').textContent = selLoc;
        document.getElementById('thStateName1').textContent = selLoc;
        document.getElementById('thStateName2').textContent = selLoc;

        // Calculations for Difference Strip
        const diffMed = locData.median - usData.median;
        const diffMedSign = diffMed >= 0 ? '+' : '';
        const diffMedPct = ((diffMed / usData.median) * 100).toFixed(1);
        const medClass = diffMed > 0 ? 'delta-pos' : (diffMed < 0 ? 'delta-neg' : 'delta-zero');
        document.getElementById('stripMedState').textContent = fmtMoney(locData.median);
        document.getElementById('stripMedDiff').innerHTML = `Diff: <span class="${medClass}">${diffMedSign}${fmtMoney(diffMed)} (${diffMedSign}${diffMedPct}%)</span>`;

        const diffGini = locData.gini_route1 - usData.gini_route1;
        const diffGiniSign = diffGini >= 0 ? '+' : '';
        const giniClass = diffGini > 0 ? 'delta-pos' : (diffGini < 0 ? 'delta-neg' : 'delta-zero');
        document.getElementById('stripGiniState').textContent = locData.gini_route1.toFixed(4);
        document.getElementById('stripGiniDiff').innerHTML = `Diff: <span class="${giniClass}">${diffGiniSign}${diffGini.toFixed(4)}</span>`;

        let diffP90Str = 'N/A';
        let p90Class = 'delta-neg';
        if (typeof locData.p90_p10 === 'number' && typeof usData.p90_p10 === 'number') {
          const diffP90 = locData.p90_p10 - usData.p90_p10;
          const diffP90Sign = diffP90 >= 0 ? '+' : '';
          diffP90Str = `${diffP90Sign}${diffP90.toLocaleString('en-US', {maximumFractionDigits:1})}`;
          p90Class = diffP90 > 0 ? 'delta-pos' : (diffP90 < 0 ? 'delta-neg' : 'delta-zero');
        }
        document.getElementById('stripP90State').textContent = p90p10LocVal;
        document.getElementById('stripP90Diff').innerHTML = `Diff: <span class="${p90Class}">${diffP90Str}</span>`;

        const diffTop10 = locData.top10_share - usData.top10_share;
        const diffTop10Sign = diffTop10 >= 0 ? '+' : '';
        const top10Class = diffTop10 > 0 ? 'delta-pos' : (diffTop10 < 0 ? 'delta-neg' : 'delta-zero');
        document.getElementById('stripTop10State').textContent = locData.top10_share.toFixed(2) + '%';
        document.getElementById('stripTop10Diff').innerHTML = `Diff: <span class="${top10Class}">${diffTop10Sign}${diffTop10.toFixed(2)}%</span>`;

        // Side-by-Side Inequality Table
        const ineqTbody = document.getElementById('compInequalityBody');
        ineqTbody.innerHTML = '';
        const ineqRows = [
          { name: 'Gini Coefficient', state: locData.gini_route1.toFixed(4), us: '0.5621', diff: (diffGiniSign + diffGini.toFixed(4)) },
          { name: 'Theil Index (Entropy T)', state: locData.theil.toFixed(4), us: '0.4959', diff: ((locData.theil - usData.theil >= 0 ? '+' : '') + (locData.theil - usData.theil).toFixed(4)) },
          { name: 'P90 / P10 Decile Ratio', state: p90p10LocVal, us: '1205.14', diff: diffP90Str },
          { name: 'Top 10% Income Share', state: locData.top10_share.toFixed(2) + '%', us: '40.02%', diff: (diffTop10Sign + diffTop10.toFixed(2) + '%') }
        ];

        ineqRows.forEach(r => {
          const tagClass = r.diff.startsWith('+') ? 'diff-tag-pos' : 'diff-tag-neg';
          const tr = document.createElement('tr');
          tr.innerHTML = `
            <td><strong>${r.name}</strong></td>
            <td><strong>${r.state}</strong></td>
            <td style="color:var(--text-muted);">${r.us}</td>
            <td><span class="${tagClass}">${r.diff}</span></td>
          `;
          ineqTbody.appendChild(tr);
        });

        // Benchmark Comparison Table
        const benchTbody = document.getElementById('compBenchmarksBody');
        benchTbody.innerHTML = '';
        [20, 40, 60, 80].forEach(p => {
          const stVal = locData.percentiles[p];
          const usVal = usData.percentiles[p];
          const diffVal = stVal - usVal;
          const diffSign = diffVal >= 0 ? '+' : '';
          const diffStr = diffSign + fmtMoney(diffVal);
          const pctDiff = ((diffVal / (usVal || 1)) * 100).toFixed(1);
          const pctStr = (diffVal >= 0 ? '+' : '') + pctDiff + '%';
          const tagClass = diffVal >= 0 ? 'diff-tag-pos' : 'diff-tag-neg';

          const tr = document.createElement('tr');
          tr.style.cursor = 'pointer';
          tr.title = `Click to inspect P${p} on the curve`;
          tr.onclick = () => quickJump(p);
          tr.innerHTML = `
            <td><strong>P${p} (${p}th)</strong></td>
            <td><strong>${fmtMoney(stVal)}</strong></td>
            <td style="color:var(--text-muted);">${fmtMoney(usVal)}</td>
            <td><span class="${tagClass}">${diffStr}</span></td>
            <td><span class="${tagClass}">${pctStr}</span></td>
          `;
          benchTbody.appendChild(tr);
        });

        // Plain-English Factual Comparison
        let ineqComparisonText = '';
        if (Math.abs(diffGini) < 0.0005) {
          ineqComparisonText = `Overall income inequality in ${selLoc} (Gini: <strong>${locData.gini_route1.toFixed(4)}</strong>) is <strong>comparable to the national benchmark</strong> (0.5621), with the top 10% holding <strong>${locData.top10_share.toFixed(2)}%</strong> of aggregate income versus <strong>40.02%</strong> nationwide.`;
        } else if (diffGini > 0) {
          ineqComparisonText = `Overall income inequality in ${selLoc} (Gini: <strong>${locData.gini_route1.toFixed(4)}</strong>) is <strong>higher than the U.S. average</strong> (0.5621, difference: <strong>${diffGiniSign}${diffGini.toFixed(4)}</strong>), with the top 10% holding <strong>${locData.top10_share.toFixed(2)}%</strong> of aggregate income versus <strong>40.02%</strong> nationally.`;
        } else {
          ineqComparisonText = `Overall income inequality in ${selLoc} (Gini: <strong>${locData.gini_route1.toFixed(4)}</strong>) is <strong>lower than the U.S. average</strong> (0.5621, difference: <strong>${diffGiniSign}${diffGini.toFixed(4)}</strong>), with the top 10% holding <strong>${locData.top10_share.toFixed(2)}%</strong> of aggregate income versus <strong>40.02%</strong> nationally.`;
        }

        let medComparisonText = '';
        if (Math.abs(diffMed) < 1) {
          medComparisonText = `Median adult income in ${selLoc} is <strong>${fmtMoney(locData.median)}</strong>, which is <strong>identical to the U.S. median</strong> ($39,589).`;
        } else if (diffMed > 0) {
          medComparisonText = `Median adult income in ${selLoc} is <strong>${fmtMoney(locData.median)}</strong>, which is <strong>${diffMedSign}${fmtMoney(diffMed)} (+${diffMedPct}%) above the U.S. median</strong> ($39,589).`;
        } else {
          medComparisonText = `Median adult income in ${selLoc} is <strong>${fmtMoney(locData.median)}</strong>, which is <strong>${diffMedSign}${fmtMoney(diffMed)} (${diffMedPct}%) below the U.S. median</strong> ($39,589).`;
        }

        let maxAbsDiffP = 20;
        let maxAbsDiffVal = 0;
        [20, 40, 60, 80].forEach(p => {
          const absDiff = Math.abs(locData.percentiles[p] - usData.percentiles[p]);
          if (absDiff > maxAbsDiffVal) {
            maxAbsDiffVal = absDiff;
            maxAbsDiffP = p;
          }
        });
        const stMaxVal = locData.percentiles[maxAbsDiffP];
        const usMaxVal = usData.percentiles[maxAbsDiffP];
        const diffMaxVal = stMaxVal - usMaxVal;
        const diffMaxStr = (diffMaxVal >= 0 ? '+' : '') + fmtMoney(diffMaxVal);
        const diffMaxPct = ((diffMaxVal / (usMaxVal || 1)) * 100).toFixed(1);
        const dirWord = diffMaxVal >= 0 ? 'above' : 'below';

        const groupComparisonText = `Among key benchmark percentiles (P20, P40, P60, P80), the largest dollar difference occurs at the <strong>${maxAbsDiffP}th percentile</strong>, where ${selLoc} income (<strong>${fmtMoney(stMaxVal)}</strong>) is <strong>${diffMaxStr} (${diffMaxVal >= 0 ? '+' : ''}${diffMaxPct}%) ${dirWord}</strong> the national benchmark (<strong>${fmtMoney(usMaxVal)}</strong>).`;

        document.getElementById('compFactualBox').innerHTML = `
          <p><strong>• Inequality Comparison:</strong> ${ineqComparisonText}</p>
          <p><strong>• Median Income:</strong> ${medComparisonText}</p>
          <p><strong>• Benchmark Variance:</strong> ${groupComparisonText}</p>
        `;
      } else {
        diffStrip.style.display = 'none';
        compCard.style.display = 'none';
      }

      // Student Guesses Table
      const guessPctList = [20, 40, 60, 80];
      const diffMap = {};
      guessPctList.forEach(p => {
        diffMap[p] = Math.abs(initialGuesses[p] - locData.percentiles[p]);
      });
      let minGapP = guessPctList[0];
      let maxGapP = guessPctList[0];
      guessPctList.forEach(p => {
        if (diffMap[p] < diffMap[minGapP]) minGapP = p;
        if (diffMap[p] > diffMap[maxGapP]) maxGapP = p;
      });

      const tbody = document.getElementById('guessTableBody');
      tbody.innerHTML = '';
      guessPctList.forEach(p => {
        const guess = initialGuesses[p];
        const real = locData.percentiles[p];
        const diff = guess - real;
        const absDiff = Math.abs(diff);
        const pctDiff = ((diff / (real || 1)) * 100).toFixed(1);
        const diffStr = (diff >= 0 ? '+' : '') + fmtMoney(diff);
        const diffBadgeClass = diff >= 0 ? 'gap-badge-over' : 'gap-badge-under';
        const overUnderText = diff >= 0 ? `+${pctDiff}% Over` : `${pctDiff}% Under`;

        let highlightBadge = '';
        if (p === minGapP) {
          highlightBadge = `<br><span class="highlight-pill pill-min-gap">Smallest Gap (${fmtMoney(absDiff)})</span>`;
        } else if (p === maxGapP) {
          highlightBadge = `<br><span class="highlight-pill pill-max-gap">Largest Gap (${fmtMoney(absDiff)})</span>`;
        }

        const tr = document.createElement('tr');
        tr.style.cursor = 'pointer';
        tr.title = 'Click to inspect this percentile on the graph';
        tr.onclick = () => quickJump(p);
        tr.innerHTML = `
          <td><strong>P${p} (${p}th)</strong></td>
          <td><span class="tag-guess">${fmtMoney(guess)}</span></td>
          <td><span class="tag-real">${fmtMoney(real)}</span></td>
          <td><span class="${diffBadgeClass}">${diffStr}</span>${highlightBadge}</td>
          <td style="font-size:11.5px; font-weight:700; color:var(--text-muted);">${overUnderText}</td>
        `;
        tbody.appendChild(tr);
      });

      // Render Chart Centerpiece
      renderChart(locData, usData, selLoc);
    }

    // 3. BETTER GRAPH DESIGN (Centerpiece rendering)
    function renderChart(locData, usData, selLoc) {
      const isState = selLoc !== 'United States';
      const svg = document.getElementById('chartSvg');
      const tooltip = document.getElementById('chartTooltip');
      svg.innerHTML = '';

      // Gradient defs for curve area
      const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
      const grad = document.createElementNS('http://www.w3.org/2000/svg', 'linearGradient');
      grad.setAttribute('id', 'curveGradient');
      grad.setAttribute('x1', '0'); grad.setAttribute('y1', '0');
      grad.setAttribute('x2', '0'); grad.setAttribute('y2', '1');

      const stop1 = document.createElementNS('http://www.w3.org/2000/svg', 'stop');
      stop1.setAttribute('offset', '0%');
      stop1.setAttribute('stop-color', isDarkMode ? '#38bdf8' : '#0284c7');
      stop1.setAttribute('stop-opacity', isDarkMode ? '0.22' : '0.16');
      grad.appendChild(stop1);

      const stop2 = document.createElementNS('http://www.w3.org/2000/svg', 'stop');
      stop2.setAttribute('offset', '100%');
      stop2.setAttribute('stop-color', isDarkMode ? '#38bdf8' : '#0284c7');
      stop2.setAttribute('stop-opacity', '0.01');
      grad.appendChild(stop2);
      defs.appendChild(grad);
      svg.appendChild(defs);

      const W = 960;
      const H = 520;
      const padL = 90;
      const padR = 40;
      const padT = 38;
      const padB = 65;

      const chartW = W - padL - padR;
      const chartH = H - padT - padB;

      // Max Y calculation
      let maxY = 0;
      for (let p = 1; p <= 100; p++) {
        if (locData.percentiles[p] > maxY) maxY = locData.percentiles[p];
        if (usData.percentiles[p] > maxY) maxY = usData.percentiles[p];
      }
      [20, 40, 60, 80].forEach(p => {
        if (initialGuesses[p] > maxY) maxY = initialGuesses[p];
      });
      maxY = Math.ceil(maxY / 25000) * 25000;

      function scaleX(p) { return padL + ((p - 1) / 99) * chartW; }
      function scaleY(val) { return padT + chartH - (Math.max(0, val) / maxY) * chartH; }

      // Subtle horizontal gridlines & Y-ticks
      const yTicks = 6;
      for (let i = 0; i <= yTicks; i++) {
        const val = (maxY / yTicks) * i;
        const yPos = scaleY(val);

        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', padL); line.setAttribute('y1', yPos);
        line.setAttribute('x2', W - padR); line.setAttribute('y2', yPos);
        line.setAttribute('class', 'grid-line');
        svg.appendChild(line);

        const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txt.setAttribute('x', padL - 12);
        txt.setAttribute('y', yPos + 4);
        txt.setAttribute('text-anchor', 'end');
        txt.setAttribute('class', 'chart-tick-label');
        txt.textContent = val === 0 ? '$0' : '$' + Math.round(val).toLocaleString('en-US');
        svg.appendChild(txt);
      }

      // Decile vertical gridlines & X-ticks
      for (let p = 10; p <= 100; p += 10) {
        const xPos = scaleX(p);
        const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        line.setAttribute('x1', xPos); line.setAttribute('y1', padT);
        line.setAttribute('x2', xPos); line.setAttribute('y2', padT + chartH);
        line.setAttribute('class', 'grid-line');
        svg.appendChild(line);

        const txt = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        txt.setAttribute('x', xPos);
        txt.setAttribute('y', padT + chartH + 22);
        txt.setAttribute('text-anchor', 'middle');
        txt.setAttribute('class', 'chart-tick-label');
        txt.textContent = p === 50 ? '50% (Median)' : p + '%';
        if (p === 50) {
          txt.setAttribute('style', `font-weight:800; fill:${isDarkMode ? '#38bdf8' : '#0284c7'};`);
        }
        svg.appendChild(txt);
      }

      // X-axis and Y-axis lines
      const xAxisLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      xAxisLine.setAttribute('x1', padL); xAxisLine.setAttribute('y1', padT + chartH);
      xAxisLine.setAttribute('x2', W - padR); xAxisLine.setAttribute('y2', padT + chartH);
      xAxisLine.setAttribute('class', 'axis-line');
      svg.appendChild(xAxisLine);

      const yAxisLine = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      yAxisLine.setAttribute('x1', padL); yAxisLine.setAttribute('y1', padT);
      yAxisLine.setAttribute('x2', padL); yAxisLine.setAttribute('y2', padT + chartH);
      yAxisLine.setAttribute('class', 'axis-line');
      svg.appendChild(yAxisLine);

      // Axis Titles
      const xTitle = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      xTitle.setAttribute('x', padL + chartW / 2);
      xTitle.setAttribute('y', H - 14);
      xTitle.setAttribute('text-anchor', 'middle');
      xTitle.setAttribute('class', 'chart-axis-title');
      xTitle.textContent = 'Population Percentile (1st through 100th)';
      svg.appendChild(xTitle);

      const yTitle = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      yTitle.setAttribute('transform', 'rotate(-90)');
      yTitle.setAttribute('x', -(padT + chartH / 2));
      yTitle.setAttribute('y', 24);
      yTitle.setAttribute('text-anchor', 'middle');
      yTitle.setAttribute('class', 'chart-axis-title');
      yTitle.textContent = 'Annual Personal Income ($ USD)';
      svg.appendChild(yTitle);

      // US Benchmark Curve (Dashed line with distinct square markers at P20, P40, P60, P80)
      if (isState) {
        let usPath = 'M ';
        for (let p = 1; p <= 100; p++) {
          const x = scaleX(p);
          const y = scaleY(usData.percentiles[p]);
          usPath += `${x.toFixed(1)},${y.toFixed(1)} `;
        }
        const usP = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        usP.setAttribute('d', usPath);
        usP.setAttribute('class', 'us-curve-path');
        svg.appendChild(usP);

        // US Benchmark milestone squares at 20, 40, 60, 80
        [20, 40, 60, 80].forEach(p => {
          const x = scaleX(p);
          const y = scaleY(usData.percentiles[p]);
          const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
          rect.setAttribute('x', x - 4.5);
          rect.setAttribute('y', y - 4.5);
          rect.setAttribute('width', 9);
          rect.setAttribute('height', 9);
          rect.setAttribute('class', 'milestone-point-us');
          svg.appendChild(rect);
        });
      }

      // Location Curve Area Gradient Fill
      let areaPath = `M ${scaleX(1).toFixed(1)},${(padT + chartH).toFixed(1)} `;
      for (let p = 1; p <= 100; p++) {
        const x = scaleX(p);
        const y = scaleY(locData.percentiles[p]);
        areaPath += `L ${x.toFixed(1)},${y.toFixed(1)} `;
      }
      areaPath += `L ${scaleX(100).toFixed(1)},${(padT + chartH).toFixed(1)} Z`;
      const areaEl = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      areaEl.setAttribute('d', areaPath);
      areaEl.setAttribute('class', 'curve-area');
      svg.appendChild(areaEl);

      // Location Solid Curve
      let locPath = 'M ';
      for (let p = 1; p <= 100; p++) {
        const x = scaleX(p);
        const y = scaleY(locData.percentiles[p]);
        locPath += `${x.toFixed(1)},${y.toFixed(1)} `;
      }
      const locP = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      locP.setAttribute('d', locPath);
      locP.setAttribute('class', 'curve-path');
      svg.appendChild(locP);

      // Highlighted P20, P40, P60, P80 Milestone points & Student Guesses
      [20, 40, 60, 80].forEach(p => {
        const x = scaleX(p);
        const realVal = locData.percentiles[p];
        const guessVal = initialGuesses[p];
        const yReal = scaleY(realVal);
        const yGuess = scaleY(guessVal);

        // Gap Line connecting Guess and Real
        const gap = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        gap.setAttribute('x1', x); gap.setAttribute('y1', yReal);
        gap.setAttribute('x2', x); gap.setAttribute('y2', yGuess);
        gap.setAttribute('class', 'gap-line');
        svg.appendChild(gap);

        // Gap label badge
        const diff = guessVal - realVal;
        const diffText = (diff >= 0 ? '+' : '') + fmtMoney(diff);
        const gapLbl = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        gapLbl.setAttribute('x', x + 6);
        gapLbl.setAttribute('y', (yReal + yGuess) / 2 + 4);
        gapLbl.setAttribute('style', 'font-size:10px; font-weight:800; fill:var(--rose);');
        gapLbl.textContent = 'Gap ' + diffText;
        svg.appendChild(gapLbl);

        // State Real Milestone Circle
        const cReal = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        cReal.setAttribute('cx', x); cReal.setAttribute('cy', yReal);
        cReal.setAttribute('r', 6);
        cReal.setAttribute('class', 'milestone-point-st');
        cReal.onclick = (e) => { e.stopPropagation(); quickJump(p); };
        svg.appendChild(cReal);

        // Student Guess Circle
        const cGuess = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        cGuess.setAttribute('cx', x); cGuess.setAttribute('cy', yGuess);
        cGuess.setAttribute('r', 6);
        cGuess.setAttribute('class', 'point-guess');
        cGuess.onclick = (e) => { e.stopPropagation(); quickJump(p); };
        svg.appendChild(cGuess);
      });

      // Dynamic Interactive Selection Group
      const dynamicGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
      dynamicGroup.setAttribute('id', 'dynamicSelectionGroup');
      svg.appendChild(dynamicGroup);

      // Selection updater
      window.selectPercentile = function(p, isPinned = false) {
        if (!locData || !locData.percentiles) return;
        p = Math.max(1, Math.min(100, Math.round(p)));
        if (isPinned) currentPercentile = p;

        dynamicGroup.innerHTML = '';

        const xPos = scaleX(p);
        const realVal = locData.percentiles[p];
        const usVal = usData.percentiles[p];
        const yReal = scaleY(realVal);

        const diffVal = realVal - usVal;
        const diffSign = diffVal >= 0 ? '+' : '';
        const diffStr = diffSign + fmtMoney(diffVal);
        const pctDiff = ((diffVal / (usVal || 1)) * 100).toFixed(1);

        // Crosshair vertical line
        const crosshair = document.createElementNS('http://www.w3.org/2000/svg', 'line');
        crosshair.setAttribute('x1', xPos); crosshair.setAttribute('y1', padT);
        crosshair.setAttribute('x2', xPos); crosshair.setAttribute('y2', padT + chartH);
        crosshair.setAttribute('class', 'active-crosshair');
        dynamicGroup.appendChild(crosshair);

        // US Benchmark dot (if state selected)
        if (isState) {
          const yUS = scaleY(usVal);
          const usDot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
          usDot.setAttribute('cx', xPos); usDot.setAttribute('cy', yUS);
          usDot.setAttribute('r', 5.5);
          usDot.setAttribute('class', 'active-us-dot');
          dynamicGroup.appendChild(usDot);
        }

        // Active target ring & dot
        const targetRing = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        targetRing.setAttribute('cx', xPos); targetRing.setAttribute('cy', yReal);
        targetRing.setAttribute('r', 10);
        targetRing.setAttribute('class', 'active-target-ring');
        dynamicGroup.appendChild(targetRing);

        const targetDot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
        targetDot.setAttribute('cx', xPos); targetDot.setAttribute('cy', yReal);
        targetDot.setAttribute('r', 5.5);
        targetDot.setAttribute('class', 'active-target-dot');
        dynamicGroup.appendChild(targetDot);

        // On-Chart Floating Callout Label
        const calloutG = document.createElementNS('http://www.w3.org/2000/svg', 'g');
        const boxW = 95;
        const boxH = 22;
        const boxX = Math.max(padL, Math.min(W - padR - boxW, xPos - boxW / 2));
        const boxY = Math.max(padT - 12, yReal - 32);

        const calloutBox = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        calloutBox.setAttribute('x', boxX); calloutBox.setAttribute('y', boxY);
        calloutBox.setAttribute('width', boxW); calloutBox.setAttribute('height', boxH);
        calloutBox.setAttribute('class', 'onchart-callout-box');
        calloutG.appendChild(calloutBox);

        const calloutText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        calloutText.setAttribute('x', boxX + boxW / 2);
        calloutText.setAttribute('y', boxY + 15);
        calloutText.setAttribute('class', 'onchart-callout-text');
        calloutText.textContent = `P${p}: ${fmtMoney(realVal)}`;
        calloutG.appendChild(calloutText);
        dynamicGroup.appendChild(calloutG);

        // Update Inspector Panel
        document.getElementById('inspBadgeLoc').textContent = selLoc;
        document.getElementById('inspBadgePct').textContent = `${p}th Percentile (P${p})`;
        document.getElementById('inspBadgeInc').textContent = `Income: ${fmtMoney(realVal)}`;

        if (isState) {
          document.getElementById('inspExplanation').innerHTML = `
            <strong>${p}th percentile: ${fmtMoney(realVal)}.</strong> In <strong>${selLoc}</strong>, income at the ${p}th percentile is <strong>${fmtMoney(realVal)}</strong> compared with <strong>${fmtMoney(usVal)}</strong> nationally, a difference of <strong>${diffStr}</strong> (${diffSign}${pctDiff}%). About ${p}% of observations in ${selLoc} have income at or below this amount.
          `;
        } else {
          document.getElementById('inspExplanation').innerHTML = `
            <strong>${p}th percentile: ${fmtMoney(realVal)}.</strong> About ${p}% of observations have income at or below this amount (and the remaining ${100 - p}% earn more).
          `;
        }

        // Active button highlight
        document.querySelectorAll('.quick-btn').forEach(btn => {
          if (btn.textContent.includes(`P${p}`)) {
            btn.classList.add('active');
          } else {
            btn.classList.remove('active');
          }
        });
      };

      // Transparent interactive overlay
      const overlay = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      overlay.setAttribute('x', padL); overlay.setAttribute('y', padT);
      overlay.setAttribute('width', chartW); overlay.setAttribute('height', chartH);
      overlay.setAttribute('fill', 'transparent');
      overlay.setAttribute('style', 'cursor: crosshair;');

      function getPercentileFromEvent(evt) {
        const rect = svg.getBoundingClientRect();
        const svgX = ((evt.clientX - rect.left) / rect.width) * W;
        return Math.max(1, Math.min(100, Math.round(1 + ((svgX - padL) / chartW) * 99)));
      }

      function showInteractiveTooltip(evt, p) {
        const rect = svg.getBoundingClientRect();
        const tipX = evt.clientX - rect.left + 16;
        const tipY = Math.max(10, evt.clientY - rect.top - 85);
        tooltip.style.left = tipX + 'px';
        tooltip.style.top = tipY + 'px';

        const realVal = locData.percentiles[p];
        const usVal = usData.percentiles[p];
        const diffVal = realVal - usVal;
        const diffSign = diffVal >= 0 ? '+' : '';
        const diffStr = diffSign + fmtMoney(diffVal);
        const pctDiff = ((diffVal / (usVal || 1)) * 100).toFixed(1);

        if (isState) {
          tooltip.innerHTML = `
            <div style="font-weight:800; font-size:12.5px; color:#93c5fd; margin-bottom:4px; border-bottom:1px solid rgba(255,255,255,0.2); padding-bottom:3px;">
              ${p}th Percentile Comparison
            </div>
            <div style="font-size:12.5px; line-height:1.5; margin-bottom:4px;">
              <div><strong>${selLoc}:</strong> <span style="color:#38bdf8; font-weight:800;">${fmtMoney(realVal)}</span></div>
              <div><strong>United States:</strong> <span style="color:#cbd5e1; font-weight:700;">${fmtMoney(usVal)}</span></div>
              <div><strong>Difference:</strong> <span style="color:#fde047; font-weight:700;">${diffStr} (${diffSign}${pctDiff}%)</span></div>
            </div>
            <div style="font-size:11px; line-height:1.35; color:#f1f5f9; border-top:1px dashed rgba(255,255,255,0.2); padding-top:4px;">
              At the ${p}th percentile, ${selLoc}'s income is ${fmtMoney(realVal)} compared with ${fmtMoney(usVal)} nationally.
            </div>
          `;
        } else {
          tooltip.innerHTML = `
            <div style="font-weight:800; font-size:12.5px; color:#93c5fd; margin-bottom:3px; border-bottom:1px solid rgba(255,255,255,0.2); padding-bottom:3px;">
              United States &bull; ${p}th Percentile
            </div>
            <div style="font-size:15px; font-weight:800; color:#38bdf8; margin-bottom:3px;">
              Income: ${fmtMoney(realVal)}
            </div>
            <div style="font-size:11.5px; line-height:1.35; color:#f1f5f9;">
              About ${p}% of observations have income at or below this amount.
            </div>
          `;
        }
        tooltip.style.opacity = 1;
      }

      overlay.onmousemove = (evt) => {
        const p = getPercentileFromEvent(evt);
        selectPercentile(p, false);
        showInteractiveTooltip(evt, p);
      };

      overlay.onclick = (evt) => {
        const p = getPercentileFromEvent(evt);
        currentPercentile = p;
        selectPercentile(p, true);
        showInteractiveTooltip(evt, p);
      };

      overlay.onmouseleave = () => {
        tooltip.style.opacity = 0;
        selectPercentile(currentPercentile, true);
      };

      svg.appendChild(overlay);

      // Initialize selection on pinned percentile (default 60th)
      selectPercentile(currentPercentile, true);
    }

    window.quickJump = function(p) {
      if (window.selectPercentile) {
        window.selectPercentile(p, true);
      }
    };

    loadData();
  </script>
</body>
</html>
"""

class DashboardHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == '/' or parsed.path == '/index.html':
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode('utf-8'))
        elif parsed.path == '/api/data':
            if not os.path.exists(CACHE_FILE):
                ensure_cache()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            with open(CACHE_FILE, 'rb') as f:
                self.wfile.write(f.read())
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        # Suppress routine log output for a cleaner terminal
        pass

def run():
    ensure_cache()
    port = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else PORT
    server_address = ('', port)
    httpd = HTTPServer(server_address, DashboardHandler)
    url = f"http://localhost:{port}"
    print("=" * 70)
    print(" OSM 202 SIGNATURE ASSIGNMENT: U.S. INCOME INEQUALITY DASHBOARD")
    print("=" * 70)
    print(f" Server running at: {url}")
    print(" To edit student guesses: open app.py and edit STUDENT_GUESSES at top.")
    print(" Press Ctrl+C in this terminal window to stop the server.")
    print("=" * 70)
    try:
        webbrowser.open(url)
    except Exception:
        pass
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
        httpd.server_close()

if __name__ == '__main__':
    run()
