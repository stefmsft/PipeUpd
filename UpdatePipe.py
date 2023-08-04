from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers
from datetime import date,datetime
import pandas as pd
import openpyxl
import glob
import os
import warnings
from dotenv import load_dotenv

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

load_dotenv()

DIRECTORY_PIPE_RAW = os.getenv("DIRECTORY_PIPE_RAW")
INPUT_SUIVI_RAW = os.getenv("INPUT_SUIVI_RAW")
OUTPUT_SUIVI_RAW = os.getenv("OUTPUT_SUIVI_RAW")

SKIP_ROW = os.getenv("SKIP_ROW")

if (SKIP_ROW == None):
    SKIP_ROW=0
else:
    SKIP_ROW = int(SKIP_ROW)

GRANULARITE = os.getenv("GRANULARITE")
if (GRANULARITE == None): GRANULARITE='Date'

GRANULARITE_COL = os.getenv("GRANULARITE_COL")
if (GRANULARITE_COL == None):
    GRANULARITE_COL=0
else:
    GRANULARITE_COL = int(GRANULARITE_COL)



################################################################
# Functions Helper
################################################################


def GetLatestPipe(idir):

    files = glob.glob(f'{idir}/*.xls*')
    latest_file = max(files, key=os.path.getctime)

    return(latest_file)

#Generic Mapping Functions
def Mapping_Generic  (Key,Col):
    rtv = ''

    try:
        rowval = df_master.loc[df_master['Key'] == Key]
        if (len(rowval) != 0):
            rtv = rowval.at[rowval.index[-1],Col]
            if (rtv == None):
                rtv = ''
    except:
        pass

    return rtv

#Mapping Functions for
# 'Estimated\nQuantity', 'Revenu From\nEstinated Qty', 'Quarter Invoice\nFacturation', 'Forecast projet\nMenu déroulant', 'Next Step & Support demandé / Commentaire'

def Mapping_Qty (Key):

    # Rules :
    # if a '=' if found, meaning a ref to another cell, I replace this ref with the value of the cell 'Qauntité'

    eq = Mapping_Generic(Key,'Estimated\nQuantity')

    if str(eq).startswith('='):
        rowval = df_master.loc[df_master['Key'] == Key]
        eq = rowval['Quantité'].values[0]

    return eq

def Mapping_RevEur (Key):

    # Rules :
    # if the cell is not empty, it has either a value or a ref to another cell (start with '=')
    # if it's a ref ... I replace this ref with the only acceptable value for the cell : 'Prix total'
    # if not, I fill the cell with the result of the Estimated Quantity multiplied by the 'Prix de vente'

    re = Mapping_Generic(Key,'Revenu From\nEstinated Qty')

    try:

        if re != None:
            if re != '':
                rowval = df_master.loc[df_master['Key'] == Key]
                if str(re).startswith('='):
                    re = rowval['Prix total'].values[0]
                else:
                    re = rowval['Estimated\nQuantity'].values[0] * rowval['Prix de vente'].values[0]
    except:
        pass


    return re

def Mapping_QtrInvoice (Key):

    return Mapping_Generic(Key,'Quarter Invoice\nFacturation')

def Mapping_FrCast (Key):

    return Mapping_Generic(Key,'Forecast projet\nMenu déroulant')

def Mapping_NxtStp (Key):

    return Mapping_Generic(Key,'Next Step & Support demandé / Commentaire')

def Format_Cell(WS,start,ColIdx,Format):

    for r in range(start,WS.max_row+1):
        WS.cell(r,ColIdx).number_format = Format

    return

def Write2Log(wb,DataLst):

    # Check if "Pipe Log" is present
    shl = wb.sheetnames
    if "Pipe Log" not in shl:
        wslog = wb.create_sheet("Pipe Log")
        Flag = True
    else:
        wslog = wb['Pipe Log']
        Flag = False

    last_row = wslog.max_row

    if last_row == 1:
        df_log = pd.DataFrame(columns = ['Date', 'WK', 'Nb OPTY','Sales Force Amount','Estimated Amount'])
    else:
        df_log = pd.DataFrame(wslog.values)
        # Set column name from new first row
        df_log.columns = df_log.iloc[0]
        # Reset the Index
        df_log.drop(df_log.index[0],inplace=True)

    rowval = df_log.loc[df_log[GRANULARITE] == DataLst[GRANULARITE_COL]]
    if (len(rowval) != 0):
        idx = df_log.index[df_log[GRANULARITE] == DataLst[GRANULARITE_COL]]
        df_log.loc[idx] = DataLst
    else:
        df_log.loc[len(df_log)+1] = DataLst

    wslog.delete_rows(2, amount=(wslog.max_row - 1))

    for r in dataframe_to_rows(df_log, index=False, header=Flag):
        wslog.append(r)

    Format_Cell(wslog,2,1,numbers.FORMAT_DATE_DDMMYY)
    Format_Cell(wslog,2,4,'[$EUR ]#,##0_-')
    Format_Cell(wslog,2,5,'[$EUR ]#,##0_-')

    return


def main():

    global df_master

    LatestPipe = GetLatestPipe(DIRECTORY_PIPE_RAW)

    ####################################
    # Load Latest Pipe File
    ####################################

    print(f'- Utilisation du fichier pipe : {LatestPipe}')
    # Skip SKIP_ROW if extract made with header details. Depending on the header lines this value can be updated from .env file
    df_pipe = pd.read_excel(LatestPipe, skiprows=SKIP_ROW)

    #Drop Empty Columns
    for i in df_pipe.columns:
        if i.startswith('Unnamed:'):
            df_pipe = df_pipe.drop(i, axis=1)

    # Reorg Columns to fit the expected Master Format
    # 'Opportunity Owner','Created Date','Close Date','Stage','Opportunity Number','Indirect Account','End Customer','Estimated Quantity','Sales Price','Estimated Total Price','Sales Model Name','Part Number','Account Name','Product Line','Deal Type'
    cols = list(df_pipe.columns.values)
    # Col numbers starts at 0
    Cval = cols.pop(11) 
    cols.insert(7, Cval)
    Cval = cols.pop(11) 
    cols.insert(7, Cval)
    df_pipe = df_pipe.reindex(columns=cols)

    ####################################
    # Cleanup Data
    ####################################

    # Bogus Values
    # 'Total', nan, 'Confidential Information - Do Not Distribute', 'Copyright © 2000-2023 salesforce.com, inc. All rights reserved.'
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Total'].index, inplace=True)
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Confidential Information - Do Not Distribute'].index, inplace=True)
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Copyright © 2000-2023 salesforce.com, inc. All rights reserved.'].index, inplace=True)
    df_pipe.dropna(subset=['Opportunity Owner'], inplace=True)


    print(f'  - Il contient {len(df_pipe)} lignes')

    # Owner to keep
    # 'William ROMAN', 'Corinne CORDEIRO', 'Kajanan SHAN', 'Younes Giaccheri', 'Aziz ABELHAOU', 'Hippolyte FOVIAUX', 'Hatem ABBACI', 'Mehdi Dahbi', 'Gwenael BOJU', 'Charles TEZENAS', Etc ...
    
    # Owner to drop ??
    # 'Clement VIEILLEFONT', 'Vincent HALLER', 'Mathieu LUTZ'
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Clement VIEILLEFONT'].index, inplace=True)
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Vincent HALLER'].index, inplace=True)
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Mathieu LUTZ'].index, inplace=True)

    # Remove "Run Rate" Type  Deals
    df_pipe.drop(df_pipe.loc[df_pipe['Deal Type']=='Run Rate Deal'].index, inplace=True)

    # Cleanup OPTY (remove NaN)
    df_pipe['Opportunity Number'].fillna("", inplace=True)
    df_pipe['Sales Model Name'].fillna("", inplace=True)

    #Format Dates
    df_pipe['Created Date'] = df_pipe['Created Date'].apply(pd.to_datetime, format='mixed')
    df_pipe['Close Date'] = df_pipe['Close Date'].apply(pd.to_datetime, format='mixed')

    # Create Key Columns (Opty+Model)
    df_pipe['Key'] = df_pipe.apply(lambda row: f'{row["Opportunity Number"]}{row["Sales Model Name"]}', axis = 1)

    print(f'  - {len(df_pipe)} lignes apres nettoyage')

    ####################################
    # Load PipeLine Excel File and convert the 'Pipeline Sell Out' Tab to DataFrame
    ####################################

    myworkbook=openpyxl.load_workbook(INPUT_SUIVI_RAW, keep_vba=True)
    worksheet= myworkbook['Pipeline Sell Out']

    df_master = pd.DataFrame(worksheet.values)

    print(f'- Selection de l onglet Pipe Sell Out du fichier {INPUT_SUIVI_RAW}')


    # Drop first row
    df_master.drop(index=df_master.index[0], axis=0, inplace=True)

    # Set column name from new first row
    df_master.columns = df_master.iloc[0]
    # Reset the Index
    df_master = df_master.reset_index(drop=True)

    print(f'  - Il contient {len(df_master) - 1} lignes')
    print(f'- Injection / refresh des dernieres OPTY ...')

    ####################################
    # Clean and Copy the previous value of updated columns from df_master in df_pipe when the corresponding Key (OPTY+MODEL) match
    ####################################

    # Columns de df_master
    # Common with Pipe File
    # Old Format : 'Propriétaire de l'opportunité', 'Date de création','Date de clôture', 'Etape', 'Opportunity Number', 'Revendeur','Client Final', 'Quantité', 'Prix de vente', 'Prix total','Nom du produit', 'Code du produit', 'Grossiste','Product Family\n(NX, NB, NR, PD, PT, PF)','Category Deal\nStock /CTO'
    # 'Propriétaire de l'opportunité', 'Opportunity Number', 'Date de création', 'Date de clôture', 'Étape', 'Revendeur', 'Client Final', 'Nom du produit', 'Code du produit', 'Quantité', 'Prix de vente', 'Prix total', 'Grossiste',
    # Added for Sales manual change (5 cols)
    # 'Qty\nUnit', 'Revenu projet\nK Euros', 'Quarter Invoice\nFacturation', 'Forecast projet\nMenu déroulant', 'Next Step & Support demandé / Commentaire'
 
    # Cleanup OPTY and Model Name (remove NaN)
    df_master['Opportunity Number'].fillna("", inplace=True)
    df_master['Nom du produit'].fillna("", inplace=True)

    # Create Key Columns (Opty+Model)
    df_master['Key'] = df_master.apply(lambda row: f'{row["Opportunity Number"]}{row["Nom du produit"]}', axis = 1)
    # Master columns used for the Key while transitioning Columns Names
    #df_master['Key'] = df_master.apply(lambda row: f'{row["Date de création"]}{row["Quantité"]}', axis = 1)

    # Column Quantity
    df_pipe['Estimated\nQuantity'] = df_pipe['Key'].map(Mapping_Qty)

    # Column Revenu projet
    df_pipe['Revenu From\nEstinated Qty'] = df_pipe['Key'].map(Mapping_RevEur)

    # Column Quarter Invoice
    df_pipe['Quarter Invoice\nFacturation'] = df_pipe['Key'].map(Mapping_QtrInvoice)

    # Column Forecast projet
    df_pipe['Forecast projet\nMenu déroulant'] = df_pipe['Key'].map(Mapping_FrCast)

    # Column Next Step
    df_pipe['Next Step & Support demandé / Commentaire'] = df_pipe['Key'].map(Mapping_NxtStp)

    # Remove "Étape:Rejected"
    df_pipe.drop(df_pipe.loc[df_pipe['Stage']=='Rejected'].index, inplace=True)

    # No need of the Key Column anymore
    df_pipe.drop(['Key'], axis=1, inplace=True)
    df_master.drop(['Key'], axis=1, inplace=True)

    SFPipeAmmount = df_pipe['Estimated Total Price'].sum()
    EstPipeAmmount = df_pipe['Revenu From\nEstinated Qty'].apply(pd.to_numeric, errors='coerce').sum()

    df_pipe.columns = df_master.columns

    worksheet.delete_rows(3, amount=(worksheet.max_row - 2))

    for r in dataframe_to_rows(df_pipe, index=False, header=False):
        worksheet.append(r)

    print(f'  - l onglet contient {len(df_pipe)} lignes maintenant')

    # Apply Columns Formats
    # Col C = 2
    Format_Cell(worksheet,3,2,numbers.FORMAT_DATE_DDMMYY)
    # Col C = 3
    Format_Cell(worksheet,3,3,numbers.FORMAT_DATE_DDMMYY)

    # Col K = 9
    Format_Cell(worksheet,3,9,numbers.FORMAT_CURRENCY_EUR_SIMPLE)
    # Col L = 10
    Format_Cell(worksheet,3,10,'[$EUR ]#,##0_-')
    # Col Q = 17
    Format_Cell(worksheet,3,17,'[$EUR ]#,##0_-')

    # Log Pipe Data
    lst = [datetime(datetime.now().year,datetime.now().month,datetime.now().day,0,0), date.today().isocalendar()[1], worksheet.max_row - 2, SFPipeAmmount, EstPipeAmmount]
    Write2Log(myworkbook,lst)

    myworkbook.save(OUTPUT_SUIVI_RAW)

    print(f'- Sauvegarde vers {OUTPUT_SUIVI_RAW}')


if __name__ == "__main__":
    main()