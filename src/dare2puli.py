#!/usr/bin/env python

from csv import DictReader
from datetime import datetime
from pathlib import Path
from tkinter import Tk
from tkinter.filedialog import askopenfilename

import tomllib
from numpy import nan
from pandas import DataFrame, read_excel, to_datetime

from utils import to_excel_table


def load_config() -> dict:
    """Returns a dict with the configuration: a) Loaded from `mock_mapping.csv`,
    and b) from `config.toml`.
    """
    config_path = Path('./config/mock_mapping.csv')
    with open(config_path, mode='r', encoding='utf-8') as f_csv:
        csv_reader = DictReader(f_csv, delimiter=';')
        mapping = {
            next(iter(row.values())): {k: v for k, v in list(row.items())[1:]}
            for row in csv_reader
        }

    publishers = list(mapping['DOI'].keys())

    with open(Path('./config/config.toml'), mode='rb') as f_toml:
        toml = tomllib.load(f_toml)

    return {
        'publishers': publishers,
        'mapping': mapping,
        'toml': toml
    }

CFG = load_config()



def ask_filepath() -> Path:
    """Aks via `tkinter.filedialog.askopenfile` for the path to the report xlsx file,
    checks if a filepath is provided, and returns it.
    """
    title = 'dare2puli.py: Choose your report .xlsx file'
    filetypes = [('Excel Files', '*.xlsx'), ('All files', '*.*')]
    filepath = askopenfilename(title=title, filetypes=filetypes)
    if not filepath:
        raise ValueError("Need filepath. Can't continue without.")
    filepath = Path(filepath)
        
    return filepath


def publisher_from_filepath(filepath: Path, publishers: list[str]) -> str:
    """Checks for every publisher in `CFG['publishers']` if he is in the filepath,
    returns the first publisher it finds or raises an `Exception` if no publisher is in filepath.
    """

    for publisher in publishers:
        if publisher.casefold() in str(filepath).casefold().replace('-', ' '):
            return publisher
    
    raise ValueError('No publisher in filepath. Can\'t continue.')


def dare_to_puli(filepath: Path, publisher: str) -> DataFrame:
    """Reads from the provided filepath the report.
    Creates empty DataFrame with the column names of the publisher list.
    Takes the wanted columns from the report to the publisher list DataFrame.
    Cleans the publisher list and adds default values.
    """
    report = read_excel(
        filepath, 
        skiprows=CFG['toml']['skiprows'].get(publisher, None),
        skipfooter=CFG['toml']['skipfooter'].get(publisher, 0),
        dtype=str,
        engine='openpyxl'
    )

    # Are all expected columns in the dashboard report?
    expected_columns = [col[publisher] for col in CFG['mapping'].values() if (col[publisher] and not col[publisher].startswith('#'))]
    missing_columns = [col for col in expected_columns if col not in report.columns]
    if missing_columns:
        for miscol in missing_columns: 
            print(f'[-] MISSING COLUMN IN DASHBOARD-REPORT: {miscol}')

    # Create empty `DataFrame` with all columns we want the publisher list to have
    puli = DataFrame(columns=CFG['mapping'], dtype=object)

    # Take the dashboard reports columns in the publisher list
    for dare_col, puli_col in zip(CFG['mapping'].values(), CFG['mapping'].keys()):
        if not dare_col[publisher] or dare_col[publisher].startswith('#'):
            continue
        if dare_col[publisher].startswith('-->'):
            puli[puli_col] = dare_col[publisher].lstrip('->').strip()
        else:
            puli[puli_col] = report[dare_col[publisher]]

    puli['Verlag'] = publisher

    puli = publisher_specific_enrichment(puli, report, publisher)

    puli = clean_data(puli, publisher)

    return puli


def publisher_specific_enrichment(puli: DataFrame, report: DataFrame, publisher: str) -> DataFrame:
    """This function is the point where you can enrich or clean the data for a specific publisher.
    """
    match publisher:
        case 'Becher Universitäts Presse':
            puli['CorresAuthor'] = report['First Author First Name'] + ' ' + 'First Author Last Name'

        case 'Magpie':
            puli['DOI'] = report['doi'].apply(lambda doi: doi.split('doi.org/')[-1] if isinstance(doi, str) else nan)

        case _:
            pass

    return puli


def clean_data(puli: DataFrame, publisher: str):
    """Does some cleaning of some `DataFrame` columns and provides two default values.
    """

    puli['Artikeltitel'] = puli['Artikeltitel'].str.strip()

    puli['CC-Lizenz'] = (
        puli['CC-Lizenz']
        .str.replace('CC-BY', 'CC BY')
        .str.replace('CC_BY', 'CC BY')
        .str.replace(' 4.0', '')
        .str.strip()
    )

    # puli['Transformationsvertrag'] = nan
    match publisher:
        case 'Laufer':
            puli['Transformationsvertrag'] = 'AGREEMENT Laufer'
        case 'Magpie':
            puli['Transformationsvertrag'] = 'AGREEMENT Magpie'
        case 'Wonka':
            puli['Transformationsvertrag'] = 'AGREEMENT Wonka'
        case _:
            puli['Transformationsvertrag'] = nan
    
    date_columns = ['Submission Date', 'Acceptance Date', 'Approval Date', 'Published Online Date']
    for date_column in date_columns:
        puli[date_column] = to_datetime(puli[date_column], format='mixed').dt.date

    for hybrid_variant in ['Hybrid Open Access', 'Hybrid', 'Online Open']:
        puli.loc[(puli['OA-Status'] == hybrid_variant), 'OA-Status'] = 'hybrid-oa'
    for gold_variant in ['Full Open Access', 'FullyOpenAccess', 'Open Access', 'Fully Open Access', 'Fully open access']:
        puli.loc[(puli['OA-Status'] == gold_variant), 'OA-Status'] = 'gold-oa'

    return puli




def save_df_as_table(puli: DataFrame, file_path: Path) -> None:
    """Saves the DataFrame using the `to_excel_table` function from utils.py
    """

    if not file_path.parent.is_dir():
        file_path.parent.mkdir()

    to_excel_table(puli.sort_values(by='Approval Date'), file_path, index=False)
        
    print(f'[+] SAVED: {file_path} ({len(puli)} items in {puli.shape[1]} columns)')



def main() -> None:
    
    Tk().withdraw()

    filepath = ask_filepath()
    print(f'[+] Filepath: {filepath}')
    
    publisher = publisher_from_filepath(filepath, CFG['publishers'])
    print(f'[+] Publisher: {publisher}')
    
    puli = dare_to_puli(filepath, publisher)

    file_path = Path('PublisherLists') / f'publisher-list_{datetime.now().astimezone().date()}_{publisher}.xlsx'
    save_df_as_table(puli, file_path)


if __name__ == '__main__':
    main()
