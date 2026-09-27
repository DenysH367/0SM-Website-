# OSM 202 Signature Assignment: Income Inequality & Distribution Dashboard

A local Python application for analyzing and visualizing U.S. income distribution and inequality metrics using the 100,000-adult American Community Survey (ACS) sample dataset (`us_adults_sample_100k.xlsx`).

---

## Quick Start Guide

### 1. Requirements
* **Python 3.8+** (Standard library only; **0 external packages needed!** No `pip install` required).

### 2. How to Run the App
Open your Terminal, navigate to this folder, and run:

```bash
python3 app.py
```

Your web browser will automatically open:
```
http://localhost:8050
```

---

## Features

1. **Geographic Selection:**
   - United States (National benchmark)
   - All 50 U.S. states + District of Columbia
2. **Income Percentiles (1 through 100):**
   - Automatically computed for every percentile using the `income` variable (Column N).
3. **Interactive Percentile Distribution Curve & Live Inspector:**
   - X-axis: Population percentile (1 to 100)
   - Y-axis: Annual income in $ USD
   - **Click & Hover Inspection:** Hovering or clicking anywhere along the line displays a vertical crosshair, an exact percentile marker, and an explanation card showing:
     - The percentile selected ($P_k$).
     - The actual income at that percentile.
     - A simple explanation: *"About $k$% of observations have income at or below $[Amount], while the remaining $(100-k)$% earn more."*
     - The selected state or United States.
     - Benchmark comparisons against the U.S. national curve and median.
   - **Quick Jump Buttons:** Instant 1-click inspection for P10, P20, P40, P50 (Median), P60, P80, P90, and P99.
4. **Student Salary Guesses (Editable):**
   - Pre-configured at 20th, 40th, 60th, and 80th percentiles.
   - Visible directly on the graph next to actual incomes with gap lines and variance percentages.
   - To edit: open `app.py` and modify `STUDENT_GUESSES` at the top of the file:
     ```python
     STUDENT_GUESSES = {
         20: 20000,   # 20th percentile guess
         40: 38000,   # 40th percentile guess
         60: 60000,   # 60th percentile guess
         80: 95000,   # 80th percentile guess
     }
     ```
5. **Inequality Measures (Selected Location vs. United States):**
   - **Gini Coefficient:** Includes **two independent calculations**:
     - *Route 1:* Geometric numerical integration of the Lorenz curve (Trapezoidal rule).
     - *Route 2:* Relative Mean Difference (RMD) sorted rank covariance formula.
     - Both yield identical headline figures, demonstrating independent computational verification.
   - **Theil Index (Entropy T):** Measures upper-tail concentration across positive earners.
   - **P90 / P10 Decile Ratio:** Measures income dispersion between top and bottom deciles.
   - **Top 10% Income Share:** Aggregate income share captured by the top 10% of adults.
6. **State vs. United States Comparison:**
   - **Compact Difference Strip:** Instant at-a-glance horizontal card comparing Median Income, Gini, P90/P10, and Top 10% Share with factual deltas.
   - **Dual Income Curve Overlay:** Displays both state (solid line + area fill) and U.S. (dashed line + benchmark square markers) simultaneously with clear legend.
   - **Interactive Dual-Curve Hover:** Hovering over any percentile displays both amounts, difference, and plain-English explanation.
   - **Side-by-Side Inequality & Benchmark Tables:** 5-column breakdown table (P20, P40, P60, P80) with clickable rows.
   - **Objective Comparative Narrative:** Plain-English summary comparing inequality, median, and benchmark variances strictly without subjective words.
   - **National View:** Selecting "United States" hides the comparison strip and displays the clean national benchmark.
7. **Clean Metric Cards with Info Popovers:**
   - Uniform dimensions, typography, and hover `ⓘ` popovers explaining Gini, Theil, P90/P10, and Top 10% share.
8. **Optional Dark Mode & Presentation Mode:**
   - `🌙 Dark Mode` / `☀️ Light Mode` toggle for high-contrast presentation environments.
   - `📺 Presentation Mode` toggle enlarging the graph and metrics while minimizing scrolling.
9. **Smooth State Changes:**
   - Subtle 120ms crossfade transition when changing locations without jarring animations.
10. **Student Guess Flow & Gap Highlights:**
   - Clearly displays: `Student Guess → Actual Income → Dollar Difference → Over/Under`.
   - Objectively flags the `Smallest Gap` and `Largest Gap`.
11. **Built-in Presentation Guide & Footer:**
   - Methodological explanations, dual-route Gini verification notes, and source attribution.

---

## Treatment of Zero and Negative Incomes

- **Negative Incomes:** Net business/farm losses (e.g. minimum -$10,304 in this dataset) are preserved in overall sample distributions and percentiles, but truncated to zero for Gini/Theil calculations where non-negative values or logarithms are mathematically required.
- **Zero Incomes:** Approximately 10% of adults in the general population report $0 personal income (unemployed, non-working students, homemakers, retirees without wage/personal income). They are retained in percentile calculations to avoid artificially skewing percentiles upward. For the Theil index, evaluation is conducted over positive earners ($y > 0$) because $\ln(0)$ is undefined.
