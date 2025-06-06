import sys
import re
import requests
from bs4 import BeautifulSoup
import pandas as pd
from collections import defaultdict

TAG_MAP = {
    'revenue': [
        'RevenueFromContractWithCustomerExcludingAssessedTax',
        'Revenues',
        'SalesRevenueNet',
    ],
    'ebitda': [
        'EBITDA',
    ],
    'operating_loss': [
        'OperatingIncomeLoss',
    ],
    'depreciation_fixed': [
        'Depreciation',
        'DepreciationDepletionAndAmortization',
    ],
    'depreciation_rou': [
        'OperatingLeaseRightOfUseAssetDepreciationExpense',
    ],
    'amortization': [
        'AmortizationOfIntangibleAssets',
    ],
    'debt': [
        'Debt',
        'LongTermDebt',
    ],
    'cash': [
        'CashAndCashEquivalentsAtCarryingValue',
    ],
    'interest': [
        'InterestAndDebtExpense',
        'InterestExpense',
    ],
    'net_cash_ops': [
        'NetCashProvidedByUsedInOperatingActivities',
    ],
    'dividends': [
        'PaymentsOfDividends',
        'DividendsPaid',
    ],
}

METRICS_ORDER = [
    'revenue',
    'ebitda',
    'debt',
    'cash',
    'interest',
    'net_cash_ops',
    'dividends',
    'net_debt',
    'debt_ebitda',
    'ebitda_interest',
    'rcf_net_debt',
]


def clean_number(text):
    if text is None:
        return None
    text = text.strip()
    if text == '':
        return None
    text = text.replace(',', '')
    text = text.replace('(', '-')
    text = text.replace(')', '')
    m = re.match(r'^-?\d+(\.\d+)?$', text)
    if not m:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def parse_contexts(soup):
    contexts = {}
    for ctx in soup.find_all(['context', 'xbrli:context']):
        ctx_id = ctx.get('id')
        if not ctx_id:
            continue
        period = ctx.find(['period', 'xbrli:period'])
        if not period:
            continue
        end_date = period.find(['enddate', 'xbrli:enddate'])
        if not end_date:
            end_date = period.find(['instant', 'xbrli:instant'])
        if not end_date:
            continue
        year = end_date.text.strip()[:4]
        contexts[ctx_id] = year
    return contexts


def extract_metrics(soup, contexts):
    data = defaultdict(lambda: defaultdict(lambda: None))
    for tag in soup.find_all(['ix:nonfraction', 'ix:nonnumeric']):
        name = tag.get('name')
        if not name:
            continue
        context_ref = tag.get('contextref')
        if not context_ref or context_ref not in contexts:
            continue
        year = contexts[context_ref]
        for metric, names in TAG_MAP.items():
            if any(name.lower().endswith(n.lower()) for n in names):
                value = clean_number(tag.text)
                scale = tag.get('scale')
                if value is not None and scale:
                    try:
                        value *= 10 ** int(scale)
                    except ValueError:
                        pass
                if metric not in data[year] or data[year][metric] is None:
                    data[year][metric] = value
                break
    return data


def compute_derived(data):
    for year, metrics in data.items():
        # EBITDA
        if metrics.get('ebitda') is None:
            comp = [metrics.get('operating_loss'), metrics.get('depreciation_fixed'),
                    metrics.get('depreciation_rou'), metrics.get('amortization')]
            if all(v is not None for v in comp):
                metrics['ebitda'] = sum(comp)
        # Net Debt
        if metrics.get('debt') is not None and metrics.get('cash') is not None:
            metrics['net_debt'] = metrics['debt'] - metrics['cash']
        # Debt/EBITDA
        if metrics.get('debt') is not None and metrics.get('ebitda') not in (None, 0):
            metrics['debt_ebitda'] = metrics['debt'] / metrics['ebitda']
        # EBITDA/Interest
        if metrics.get('interest') not in (None, 0) and metrics.get('ebitda') is not None:
            metrics['ebitda_interest'] = metrics['ebitda'] / metrics['interest']
        # RCF/Net Debt
        if metrics.get('net_cash_ops') is not None and metrics.get('dividends') is not None and metrics.get('net_debt') not in (None, 0):
            rcf = metrics['net_cash_ops'] - metrics['dividends']
            metrics['rcf_net_debt'] = rcf / metrics['net_debt']
    return data


def metrics_to_dataframe(data):
    rows = []
    years = sorted(data.keys(), reverse=True)[:3]
    for year in years:
        row = {'year': year}
        for metric in METRICS_ORDER:
            row[metric] = data[year].get(metric)
        rows.append(row)
    df = pd.DataFrame(rows)
    return df


def metrics_to_excel(data, output_path):
    df = metrics_to_dataframe(data)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='metrics')
        # Optional sheet with raw metrics
        raw_df = pd.DataFrame(data).T
        raw_df.to_excel(writer, sheet_name='raw')


def main(url, output_path):
    r = requests.get(url)
    r.raise_for_status()
    soup = BeautifulSoup(r.content, 'lxml')
    contexts = parse_contexts(soup)
    data = extract_metrics(soup, contexts)
    data = compute_derived(data)
    metrics_to_excel(data, output_path)


if __name__ == '__main__':
    if len(sys.argv) != 3:
        print('Usage: python sec_metrics.py <ixbrl url> <output.xlsx>')
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
