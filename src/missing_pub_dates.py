from pathlib import Path
from time import sleep
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import numpy as np
import pandas as pd
import requests

PROXIES = {
    'http' : 'proxy.charite.de:8080',
    'https': 'proxy.charite.de:8080',
}

SLEEP_TIME  = 1.5
RESULTS_DIR = Path('PublisherLists')
OUTPUT_XLSX = RESULTS_DIR / 'missingPubDates.xlsx'
OUTPUT_TXT  = RESULTS_DIR / 'missingPubDates.txt'


def ask_filepath() -> str:
    """Ask via `tkinter.filedialog.askopenfile` for the path to the report xlsx file,
    checks if a filepath is provided, and returns it.
    """
    title = 'missing_pub_dates.py: Choose your xlsx file'
    filetypes=[('Excel Files', '*.xlsx'), ('All files', '*.*')]
    filepath = askopenfilename(title=title, filetypes=filetypes)
    
    if not Path(filepath).is_file():
        raise ValueError(f'{filepath} is no file.')
    
    return filepath


def api_call(url: str) -> dict:
    try:
        print(f'\t{url}')
        sleep(SLEEP_TIME)
        r = requests.get(url, proxies=PROXIES)
        assert r.status_code == requests.codes.ok, f'StatusCodeError: {r.status_code}'
        print(f'\t{r}')
        return r.json()
        
    except AssertionError as e:
        print(f'[-] ERROR: {e}')
        return {}


def parse_date(date: str):
    try:
        return pd.to_datetime(date).date()
    except:  # noqa: E722
        return date


def main():
    Tk().withdraw()
    filepath = ask_filepath()
    
    df = pd.read_excel(filepath)
    print(f'[+] DataFrame eingelesen. {df.shape[0]} rows and {df.shape[1]} columns.')

    if 'DOI' not in df.columns:
        raise ValueError('No column named "DOI" in xlsx file')
    
    if 'Online Publication Date' not in df.columns:
        df['Online Publication Date'] = np.nan
    
    print(f'[!] Column "Online Publication Date" is {df['Online Publication Date'].isna().sum()} times without entry.\n')
    
    doi_cr = {}
    for doi in df.loc[df['Online Publication Date'].isna()]['DOI'].unique():
        url = f'https://api.crossref.org/works/{doi}'
        doi_cr[doi] = api_call(url).get('message', {})
    
    doi_date = {}
    for doi, cr_j in doi_cr.items():
        try:
            dateparts = [str(d) for d in cr_j['published-online'].get('date-parts', [[]])[0]]
            date = pd.to_datetime('-'.join(dateparts)).date()
            doi_date[doi] = date
        except KeyError:
            doi_date[doi] = np.nan
    
    df.loc[
        df['Online Publication Date'].isna(), 'Online Publication Date'
    ] = df['DOI'].map(doi_date)
    
    for date_col in [col for col in df.columns if 'Date' in col]:
        # print(date_col)
        df[date_col] = df[date_col].apply(parse_date)
    
    file_out = Path(filepath).stem
    file_out_xlsx = RESULTS_DIR / f'{file_out}_withPubDates.xlsx'
    file_out_txt  = RESULTS_DIR / f'{file_out}_withPubDates.txt'
    
    with open(file_out_txt, mode='w', encoding='utf-8') as f:
        for date in df['Online Publication Date']:#.dt.date:
            if isinstance(date, float):
                f.write('\n')
            else:
                f.write(f'{date}\n')
    print(f'[+] Saved to {file_out_txt}')
    
    df.to_excel(file_out_xlsx, index=False)
    print(f'[+] Saved to {file_out_xlsx}')
    


if __name__ == '__main__':
    main()
