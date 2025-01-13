import matplotlib
import pandas as pd
matplotlib.use('Agg')
from io import BytesIO
from datetime import date
import matplotlib.pyplot as plt

df = pd.read_excel(r'C:\Users\flore.verkest\Documents\documenten\code\digiportal\digiportal\dataportal\static\data\cijfers.xlsx')
today = date.today()

def iff_g001():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    filtered_row = df.loc[df['objectnummer'] == 'IFFTOTAAL', last_12_months].squeeze()

    plt.style.use('ggplot')
    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, filtered_row, marker='o', label='Records Adlib')
    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Registratie collectie IFF')
    plt.legend()
    plt.tight_layout()
    g001 = BytesIO()
    plt.savefig(g001, format='png')
    g001.seek(0) 
    plt.close()
    
    return g001

def iff_g002():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    december_columns = [col for col in date_columns if pd.notnull(col) and col.month == 12]
    most_recent_december = max(december_columns)

    # Generate the last 12 years of January columns
    last_12_years = [
    most_recent_december.replace(year=most_recent_december.year - i)
    for i in range(12)
    ]
    
    filtered_row = df.loc[df['objectnummer'] == 'IFFTOTAAL', last_12_years].squeeze()
    plt.figure(figsize=(10, 6))
    plt.plot(last_12_years, filtered_row, marker='o', label='Records Adlib')
    plt.xticks(ticks=last_12_years, labels=[dt.strftime('%Y') for dt in last_12_years], rotation=45)
    plt.xlabel('Jaar')
    plt.ylabel('Aantal')
    plt.title('Registratie collectie IFF per jaar')
    plt.legend()
    plt.tight_layout()
    g002 = BytesIO()
    plt.savefig(g002, format='png')
    g002.seek(0) 
    plt.close()

    return g002

def iff_g003():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    filtered_rowIFF = df.loc[df['objectnummer'] == 'IFF', last_12_months].squeeze()
    filtered_rowIFFD = df.loc[df['objectnummer'] == 'IFFD', last_12_months].squeeze()
    filtered_rowIFFDA = df.loc[df['objectnummer'] == 'IFFDA', last_12_months].squeeze()
    filtered_rowIFFDC = df.loc[df['objectnummer'] == 'IFFDC', last_12_months].squeeze()
    filtered_rowIFFF = df.loc[df['objectnummer'] == 'IFFF', last_12_months].squeeze()
    filtered_rowIFFFC = df.loc[df['objectnummer'] == 'IFFFC', last_12_months].squeeze()
    filtered_rowIFFFA = df.loc[df['objectnummer'] == 'IFFFA', last_12_months].squeeze()
    filtered_rowIFFH = df.loc[df['objectnummer'] == 'IFFH', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, filtered_rowIFF, marker='o', label='IFF')
    plt.plot(last_12_months, filtered_rowIFFD, marker='o', label='IFFD')
    plt.plot(last_12_months, filtered_rowIFFDA, marker='o', label='IFFDA')
    plt.plot(last_12_months, filtered_rowIFFDC, marker='o', label='IFFDC')
    plt.plot(last_12_months, filtered_rowIFFF, marker='o', label='IFFF')
    plt.plot(last_12_months, filtered_rowIFFFC, marker='o', label='IFFFC')
    plt.plot(last_12_months, filtered_rowIFFFA, marker='o', label='IFFFA')
    plt.plot(last_12_months, filtered_rowIFFH, marker='o', label='IFFH')

    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Registratie collectie IFF')
    plt.legend()
    plt.tight_layout()
    g003 = BytesIO()
    plt.savefig(g003, format='png')
    g003.seek(0) 
    plt.close()
    return g003

############################################################################################################
def iff_g004():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    filtered_row = df.loc[df['objectnummer'] == 'BIFFTOTAAL', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, filtered_row, marker='o', label='Records Adlib met afbeelding')
    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Digitalisatie collectie IFF')
    plt.legend()
    plt.tight_layout()
    g004 = BytesIO()
    plt.savefig(g004, format='png')
    g004.seek(0) 
    plt.close()
    return g004

def iff_g005():
    date_columns = [pd.to_datetime(col, errors='coerce') for col in df.columns]
    current_month = max([col for col in date_columns if pd.notnull(col)])

    # Generate the last 12 months
    last_12_months = [
    (current_month.replace(month=(current_month.month - i - 1) % 12 + 1, 
                        year=current_month.year - (1 if current_month.month - i - 1 < 0 else 0)))
    for i in range(12)
    ]

    filtered_rowIFF = df.loc[df['objectnummer'] == 'BIFF', last_12_months].squeeze()
    filtered_rowIFFD = df.loc[df['objectnummer'] == 'BIFFD', last_12_months].squeeze()
    filtered_rowIFFDA = df.loc[df['objectnummer'] == 'BIFFDA', last_12_months].squeeze()
    filtered_rowIFFDC = df.loc[df['objectnummer'] == 'BIFFDC', last_12_months].squeeze()
    filtered_rowIFFF = df.loc[df['objectnummer'] == 'BIFFF', last_12_months].squeeze()
    filtered_rowIFFFC = df.loc[df['objectnummer'] == 'BIFFFC', last_12_months].squeeze()
    filtered_rowIFFFA = df.loc[df['objectnummer'] == 'BIFFFA', last_12_months].squeeze()
    filtered_rowIFFH = df.loc[df['objectnummer'] == 'BIFFH', last_12_months].squeeze()

    plt.figure(figsize=(10, 6))
    plt.plot(last_12_months, filtered_rowIFF, marker='o', label='IFF')
    plt.plot(last_12_months, filtered_rowIFFD, marker='o', label='IFFD')
    plt.plot(last_12_months, filtered_rowIFFDA, marker='o', label='IFFDA')
    plt.plot(last_12_months, filtered_rowIFFDC, marker='o', label='IFFDC')
    plt.plot(last_12_months, filtered_rowIFFF, marker='o', label='IFFF')
    plt.plot(last_12_months, filtered_rowIFFFC, marker='o', label='IFFFC')
    plt.plot(last_12_months, filtered_rowIFFFA, marker='o', label='IFFFA')
    plt.plot(last_12_months, filtered_rowIFFH, marker='o', label='IFFH')

    plt.xticks(ticks=last_12_months, labels=[dt.strftime('%Y-%m') for dt in last_12_months], rotation=90)
    plt.xlabel('Maand')
    plt.ylabel('Aantal')
    plt.title('Digitalisatie collectie IFF')
    plt.legend()
    plt.tight_layout()
    g005 = BytesIO()
    plt.savefig(g005, format='png')
    g005.seek(0) 
    plt.close()
    return g005

def iff_g006():
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
    g006 = BytesIO()
    plt.savefig(g006, format='png')
    g006.seek(0) 
    plt.close()
    return g006

def iff_g007():
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
    g007 = BytesIO()
    plt.savefig(g007, format='png')
    g007.seek(0) 
    plt.close()
    return g007

def iff_g008():
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
    g008 = BytesIO()
    plt.savefig(g008, format='png')
    g008.seek(0) 
    plt.close()
    return g008
