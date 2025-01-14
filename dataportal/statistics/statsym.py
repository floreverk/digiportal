import matplotlib
import pandas as pd
matplotlib.use('Agg')
from io import BytesIO
from datetime import date
import matplotlib.pyplot as plt

df = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\cijfers.xlsx')
today = date.today()

def ym_g001():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    filtered_row = df.loc[df['objectnummer'] == 'YMTOTAAL', last_12_months].squeeze()

    plt.style.use('ggplot')
    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, filtered_row, marker='o', label='Records Adlib')
    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Registratie collectie YM')
    plt.legend()
    plt.tight_layout()
    g001 = BytesIO()
    plt.savefig(g001, format='png')
    g001.seek(0) 
    plt.close()
    
    return g001

def ym_g002():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    december_columns = [col for col in date_columns if pd.notnull(col) and col.month == 12]
    most_recent_december = max(december_columns)

    # Generate the last 12 years of January columns
    last_12_years = [
    most_recent_december.replace(year=most_recent_december.year - i)
    for i in range(12)
    ]
    
    filtered_row = df.loc[df['objectnummer'] == 'YMTOTAAL', last_12_years].squeeze()
    plt.figure(figsize=(10, 6))
    plt.plot(last_12_years, filtered_row, marker='o', label='Records Adlib')
    plt.xticks(ticks=last_12_years, labels=[dt.strftime('%Y') for dt in last_12_years], rotation=45)
    plt.xlabel('Jaar')
    plt.ylabel('Aantal')
    plt.title('Registratie collectie YM per jaar')
    plt.legend()
    plt.tight_layout()
    g002 = BytesIO()
    plt.savefig(g002, format='png')
    g002.seek(0) 
    plt.close()

    return g002

def ym_g003():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    MGB = df.loc[df['objectnummer'] == 'MGB', last_12_months].squeeze()
    OM = df.loc[df['objectnummer'] == 'OM', last_12_months].squeeze()
    OMF = df.loc[df['objectnummer'] == 'OMF', last_12_months].squeeze()
    OMFA = df.loc[df['objectnummer'] == 'OMFA', last_12_months].squeeze()
    OMFC = df.loc[df['objectnummer'] == 'OMFC', last_12_months].squeeze()
    SM = df.loc[df['objectnummer'] == 'SM', last_12_months].squeeze()
    SMF = df.loc[df['objectnummer'] == 'SMF', last_12_months].squeeze()
    SMFA = df.loc[df['objectnummer'] == 'SMFA', last_12_months].squeeze()
    SMFC = df.loc[df['objectnummer'] == 'SMFC', last_12_months].squeeze()
    YM = df.loc[df['objectnummer'] == 'YM', last_12_months].squeeze()
    YMD = df.loc[df['objectnummer'] == 'YMD', last_12_months].squeeze()
    YMDA = df.loc[df['objectnummer'] == 'YMDA', last_12_months].squeeze()
    YMDC = df.loc[df['objectnummer'] == 'YMDC', last_12_months].squeeze()
    YMF = df.loc[df['objectnummer'] == 'YMF', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, MGB, marker='o', label='MGB')
    plt.plot(last_12_months, OM, marker='o', label='OM')
    plt.plot(last_12_months, OMF, marker='o', label='OMF')
    plt.plot(last_12_months, OMFA, marker='o', label='OMFA')
    plt.plot(last_12_months, OMFC, marker='o', label='OMFC')
    plt.plot(last_12_months, SM, marker='o', label='SM')
    plt.plot(last_12_months, SMF, marker='o', label='SMF')
    plt.plot(last_12_months, SMFA, marker='o', label='SMFA')
    plt.plot(last_12_months, SMFC, marker='o', label='SMFC')
    plt.plot(last_12_months, YM, marker='o', label='YM')
    plt.plot(last_12_months, YMD, marker='o', label='YMD')
    plt.plot(last_12_months, YMDA, marker='o', label='YMDA')
    plt.plot(last_12_months, YMDC, marker='o', label='YMDC')
    plt.plot(last_12_months, YMF, marker='o', label='YMF')

    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Registratie collectie YM')
    plt.legend()
    plt.tight_layout()
    g003 = BytesIO()
    plt.savefig(g003, format='png')
    g003.seek(0) 
    plt.close()
    return g003

def ym_g004():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    objectnaam = df.loc[df['objectnummer'] == 'Yobjectnaam', last_12_months].squeeze()
    titel = df.loc[df['objectnummer'] == 'Ytitel', last_12_months].squeeze()
    afmeting = df.loc[df['objectnummer'] == 'Yafmeting', last_12_months].squeeze()
    verwerving = df.loc[df['objectnummer'] == 'Yverwerving', last_12_months].squeeze()
    associatie = df.loc[df['objectnummer'] == 'Yassociatie', last_12_months].squeeze()
    iconografie = df.loc[df['objectnummer'] == 'Yiconografie', last_12_months].squeeze()
    rechten = df.loc[df['objectnummer'] == 'Yrechten', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, objectnaam, marker='o', label='objectnaam')
    plt.plot(last_12_months, titel, marker='o', label='titel')
    plt.plot(last_12_months, afmeting, marker='o', label='afmeting')
    plt.plot(last_12_months, verwerving, marker='o', label='verwerving')
    plt.plot(last_12_months, associatie, marker='o', label='associatie')
    plt.plot(last_12_months, iconografie, marker='o', label='iconografie')
    plt.plot(last_12_months, rechten, marker='o', label='rechten')

    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Basisregistratie YM')
    plt.legend()
    plt.tight_layout()
    g004 = BytesIO()
    plt.savefig(g004, format='png')
    g004.seek(0) 
    plt.close()
    return g004

############################################################################################################
def ym_g005():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    filtered_row = df.loc[df['objectnummer'] == 'YMBTOTAAL', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, filtered_row, marker='o', label='Records Adlib met afbeelding')
    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Digitalisatie collectie YM')
    plt.legend()
    plt.tight_layout()
    g005 = BytesIO()
    plt.savefig(g005, format='png')
    g005.seek(0) 
    plt.close()
    return g005

def ym_g006():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    BMGB = df.loc[df['objectnummer'] == 'BMGB', last_12_months].squeeze()
    BOM = df.loc[df['objectnummer'] == 'BOM', last_12_months].squeeze()
    BOMF = df.loc[df['objectnummer'] == 'BOMF', last_12_months].squeeze()
    BOMFA = df.loc[df['objectnummer'] == 'BOMFA', last_12_months].squeeze()
    BOMFC = df.loc[df['objectnummer'] == 'BOMFC', last_12_months].squeeze()
    BSM = df.loc[df['objectnummer'] == 'BSM', last_12_months].squeeze()
    BSMF = df.loc[df['objectnummer'] == 'BSMF', last_12_months].squeeze()
    BSMFA = df.loc[df['objectnummer'] == 'BSMFA', last_12_months].squeeze()
    BSMFC = df.loc[df['objectnummer'] == 'BSMFC', last_12_months].squeeze()
    BYM = df.loc[df['objectnummer'] == 'BYM', last_12_months].squeeze()
    BYMD = df.loc[df['objectnummer'] == 'BYMD', last_12_months].squeeze()
    BYMDA = df.loc[df['objectnummer'] == 'BYMDA', last_12_months].squeeze()
    BYMDC = df.loc[df['objectnummer'] == 'BYMDC', last_12_months].squeeze()
    BYMF = df.loc[df['objectnummer'] == 'BYMF', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, BMGB, marker='o', label='MGB')
    plt.plot(last_12_months, BOM, marker='o', label='OM')
    plt.plot(last_12_months, BOMF, marker='o', label='OMF')
    plt.plot(last_12_months, BOMFA, marker='o', label='OMFA')
    plt.plot(last_12_months, BOMFC, marker='o', label='OMFC')
    plt.plot(last_12_months, BSM, marker='o', label='SM')
    plt.plot(last_12_months, BSMF, marker='o', label='SMF')
    plt.plot(last_12_months, BSMFA, marker='o', label='SMFA')
    plt.plot(last_12_months, BSMFC, marker='o', label='SMFC')
    plt.plot(last_12_months, BYM, marker='o', label='YM')
    plt.plot(last_12_months, BYMD, marker='o', label='YMD')
    plt.plot(last_12_months, BYMDA, marker='o', label='YMDA')
    plt.plot(last_12_months, BYMDC, marker='o', label='YMDC')
    plt.plot(last_12_months, BYMF, marker='o', label='YMF')

    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Digitalisatie collectie YM')
    plt.legend()
    plt.tight_layout()
    g006 = BytesIO()
    plt.savefig(g006, format='png')
    g006.seek(0) 
    plt.close()
    return g006

def ym_g007():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    onbehandeld = df.loc[df['objectnummer'] == 'onbehandeld', last_12_months].squeeze()
    descriptor = df.loc[df['objectnummer'] == 'descriptor', last_12_months].squeeze()
    nondescriptor = df.loc[df['objectnummer'] == 'non-descriptor', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, onbehandeld, marker='o', label='te verwerken')
    plt.plot(last_12_months, descriptor, marker='o', label='descriptor')
    plt.plot(last_12_months, nondescriptor, marker='o', label='non-descriptor')

    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Thesaurus')
    plt.legend()
    plt.tight_layout()
    g007 = BytesIO()
    plt.savefig(g007, format='png')
    g007.seek(0) 
    plt.close()
    return g007

def ym_g008():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    AAT = df.loc[df['objectnummer'] == 'AAT', last_12_months].squeeze()
    TGN = df.loc[df['objectnummer'] == 'TGN', last_12_months].squeeze()
    WIKIDATA = df.loc[df['objectnummer'] == 'WIKIDATA', last_12_months].squeeze()
    IOE = df.loc[df['objectnummer'] == 'IOE', last_12_months].squeeze()
    NL = df.loc[df['objectnummer'] == 'NAMENLIJST', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, AAT, marker='o', label='AAT')
    plt.plot(last_12_months, TGN, marker='o', label='TGN')
    plt.plot(last_12_months, WIKIDATA, marker='o', label='Wikidata')
    plt.plot(last_12_months, IOE, marker='o', label='IOE')
    plt.plot(last_12_months, NL, marker='o', label='Namenlijst')

    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Thesaurus')
    plt.legend()
    plt.tight_layout()
    g008 = BytesIO()
    plt.savefig(g008, format='png')
    g008.seek(0) 
    plt.close()
    return g008

def ym_g009():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    objectnaam = df.loc[df['objectnummer'] == 'objectnaam', last_12_months].squeeze()
    periode = df.loc[df['objectnummer'] == 'periode', last_12_months].squeeze()
    geografischtrefwoord = df.loc[df['objectnummer'] == 'geografisch trefwoord', last_12_months].squeeze()
    plaats = df.loc[df['objectnummer'] == 'plaats', last_12_months].squeeze()
    materiaal = df.loc[df['objectnummer'] == 'materiaal', last_12_months].squeeze()
    techniek = df.loc[df['objectnummer'] == 'techniek', last_12_months].squeeze()
    onderwerp = df.loc[df['objectnummer'] == 'onderwerp', last_12_months].squeeze()
    rol = df.loc[df['objectnummer'] == 'rol', last_12_months].squeeze()
    gebeurtenis = df.loc[df['objectnummer'] == 'gebeurtenis', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, objectnaam, marker='o', label='objectnaam')
    plt.plot(last_12_months, periode, marker='o', label='periode')
    plt.plot(last_12_months, geografischtrefwoord, marker='o', label='geografisch trefwoord')
    plt.plot(last_12_months, plaats, marker='o', label='plaats')
    plt.plot(last_12_months, materiaal, marker='o', label='materiaal')
    plt.plot(last_12_months, techniek, marker='o', label='techniek')
    plt.plot(last_12_months, onderwerp, marker='o', label='onderwerp')
    plt.plot(last_12_months, rol, marker='o', label='rol')
    plt.plot(last_12_months, gebeurtenis, marker='o', label='gebeurtenis')

    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Thesaurus')
    plt.legend()
    plt.tight_layout()
    g009 = BytesIO()
    plt.savefig(g009, format='png')
    g009.seek(0) 
    plt.close()
    return g009
