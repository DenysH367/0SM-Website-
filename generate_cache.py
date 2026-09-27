import zipfile
import xml.etree.ElementTree as ET
import json
import math
import os

excel_path = '/Users/denys/Downloads/OSM202_Inequality/us_adults_sample_100k.xlsx'
cache_path = '/Users/denys/Downloads/OSM202_Inequality/dataset_cache.json'

print("Reading Excel...")
shared_strings = []
with zipfile.ZipFile(excel_path) as z:
    if 'xl/sharedStrings.xml' in z.namelist():
        tree = ET.parse(z.open('xl/sharedStrings.xml'))
        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        for si in tree.getroot().findall('ns:si', ns):
            shared_strings.append(''.join(t.text or '' for t in si.findall('.//ns:t', ns)))

data_by_state = {} # state_name -> list of incomes
us_incomes = []

with zipfile.ZipFile(excel_path) as z:
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
    # arr is sorted
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
    # arr is sorted
    n = len(arr)
    if n == 0:
        return {}
    
    # 1. Gini Route 1: Lorenz curve trapezoid integration (using non-negative incomes)
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
        
        # 2. Gini Route 2: Relative Mean Difference / Rank formula
        rank_sum = sum(i * y for i, y in enumerate(arr_nonneg, 1))
        gini_route2 = round((2.0 * rank_sum) / (n_nn * tot_nn) - (n_nn + 1.0) / n_nn, 4)
    
    # 3. Theil Index: T = (1/N) * sum( (y/mu) * ln(y/mu) ) for y > 0
    arr_pos = [x for x in arr if x > 0]
    n_pos = len(arr_pos)
    theil = 0.0
    if n_pos > 0:
        mu = sum(arr_pos) / n_pos
        if mu > 0:
            theil = round(sum((y / mu) * math.log(y / mu) for y in arr_pos) / n_pos, 4)
            
    # 4. P90/P10 Ratio
    p = calc_percentiles(arr)
    p10 = p[10]
    p90 = p[90]
    if p10 > 0:
        p90_p10 = round(p90 / p10, 2)
    elif p10 == 0:
        p90_p10 = "N/A (P10 = $0)"
    else:
        p90_p10 = "N/A (P10 < $0)"
        
    # 5. Top 10% Share
    # Incomes at top 10%
    k90 = int(math.floor(n * 0.90))
    top10_sum = sum(arr[k90:])
    tot_sum = sum(arr)
    if tot_sum > 0:
        top10_share = round((top10_sum / tot_sum) * 100.0, 2)
    else:
        top10_share = 0.0
        
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

print("Computing statistics for US and 51 states...")
us_incomes.sort()
output_data = {
    'United States': calc_inequality(us_incomes),
    'states': {}
}

for s_name in sorted(data_by_state.keys()):
    st_incomes = sorted(data_by_state[s_name])
    output_data['states'][s_name] = calc_inequality(st_incomes)

with open(cache_path, 'w') as f:
    json.dump(output_data, f)

print(f"Saved dataset cache to {cache_path} ({os.path.getsize(cache_path) / 1024:.1f} KB)")
