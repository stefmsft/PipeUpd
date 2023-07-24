from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers
from datetime import datetime
import pandas as pd
import openpyxl
import glob
import os
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

## Raw Data Input

DIRECTORY_PIPE_RAW="C:\\Users\\stephane_saunier\\OneDrive - ASUS\\Bizz\\SalesReports\\PipeExtracts"
#INPUT_SUIVI_RAW="C:\\Projects\\MacroBizzPipe\\Pipeline Projet - Test MACRO - Do not touch.xlsm"
INPUT_SUIVI_RAW="C:\\Users\stephane_saunier\\ASUS\\ASUS BUSINESS TEAM - Channel Weekly Meeting\\Pipeline Projet - Reunion Pipe - Updated  MACRO.xlsm"
#OUTPUT_SUIVI_RAW="C:\\Projects\\MacroBizzPipe\\Out-Test.xlsm"
OUTPUT_SUIVI_RAW="C:\\Users\stephane_saunier\\ASUS\\ASUS BUSINESS TEAM - Channel Weekly Meeting\\Pipeline Projet - Reunion Pipe - Updated  MACRO.xlsm"

def GetLatestPipe(idir):

    files = glob.glob(f'{idir}/*.xlsx')
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
# 'Category Deal\nStock /CTO', 'Product Family\n(NX, NB, NR, PD, PT, PF)', 'Qty\nUnit', 'Revenu projet\nK Euros', 'Quarter Invoice\nFacturation', 'Forecast projet\nMenu déroulant', 'Next Step & Support demandé / Commentaire'

def Mapping_CatDeal (Key):

    CapResult = Mapping_Generic(Key,'Category Deal\nStock /CTO').capitalize()

    return CapResult

def Mapping_ProdFam (Key):

    return Mapping_Generic(Key,'Product Family\n(NX, NB, NR, PD, PT, PF)')

def Mapping_Qty (Key):

    return Mapping_Generic(Key,'Qty\nUnit')

def Mapping_RevEur (Key):

    return Mapping_Generic(Key,'Revenu projet\nK Euros')

def Mapping_QtrInvoice (Key):

    return Mapping_Generic(Key,'Quarter Invoice\nFacturation')

def Mapping_FrCast (Key):

    return Mapping_Generic(Key,'Forecast projet\nMenu déroulant')

def Mapping_NxtStp (Key):

    return Mapping_Generic(Key,'Next Step & Support demandé / Commentaire')

def Format_Cell(WS,ColIdx,Format):

    for r in range(3,WS.max_row):
        WS.cell(r,ColIdx).number_format = Format

def main():

    global df_master

    LatestPipe = GetLatestPipe(DIRECTORY_PIPE_RAW)

    ####################################
    # Load Latest Pipe File
    ####################################

    print(f'- Utilisation du fichier pipe : {LatestPipe}')
    df_pipe = pd.read_excel(LatestPipe, skiprows=11)
    print(f'  - Il contient {len(df_pipe)} lignes')

    #Drop Null Columns
    df_pipe = df_pipe.drop('Unnamed: 0', axis=1)
    df_pipe = df_pipe.drop('Unnamed: 2', axis=1)

    # Reorg Columns to fit the Master Pipe Format
    # 'Opportunity Owner', 'Opportunity Number', 'Created Date', 'Close Date', 'Stage', 'Indirect Account', 'Account Name', 'Sales Model Name', 'Part Number', 'Estimated Quantity', 'Sales Price', 'Estimated Total Price', 'End Customer'
    cols = list(df_pipe.columns.values)
    Cval = cols.pop(11) 
    cols.insert(5, Cval)
    Cval = cols.pop(12) 
    cols.insert(6, Cval)
    df_pipe = df_pipe.reindex(columns=cols)

    ####################################
    # Cleanup Data
    ####################################

    # Owner to keep
    # 'William ROMAN', 'Corinne CORDEIRO', 'Kajanan SHAN', 'Younes Giaccheri', 'Aziz ABELHAOU', 'Hippolyte FOVIAUX', 'Hatem ABBACI', 'Mehdi Dahbi', 'Gwenael BOJU', 'Charles TEZENAS'
    
    # Owner to drop ??
    # 'Clement VIEILLEFONT', 'Vincent HALLER', 'Mathieu LUTZ'
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Clement VIEILLEFONT'].index, inplace=True)
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Vincent HALLER'].index, inplace=True)
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Mathieu LUTZ'].index, inplace=True)

    # Bogus Values
    # 'Total', nan, 'Confidential Information - Do Not Distribute', 'Copyright © 2000-2023 salesforce.com, inc. All rights reserved.'
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Total'].index, inplace=True)
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Confidential Information - Do Not Distribute'].index, inplace=True)
    df_pipe.drop(df_pipe.loc[df_pipe['Opportunity Owner']=='Copyright © 2000-2023 salesforce.com, inc. All rights reserved.'].index, inplace=True)
    df_pipe.dropna(subset=['Opportunity Owner'], inplace=True)

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
    print(f'  - Il contient {len(df_master) - 1} lignes')
    print(f'- Injection / refresh des dernieres OPTY ...')

    # Drop first row
    df_master.drop(index=df_master.index[0], axis=0, inplace=True)
    # Get a list on column name -> Easy to acces Columns after with df.loc[:, CName[NbCol]]
    # CName = df_master.iloc[0]
    # Set column name from new first row
    df_master.columns = df_master.iloc[0]
    # Reset the Index
    df_master = df_master.reset_index(drop=True)

    ####################################
    # Clean and Copy the previous value of updated columns from df_master in df_pipe when the corresponding Key (OPTY+MODEL) match
    ####################################

    # Columns de df_master
    # Common with Pipe File
    # 'Propriétaire de l'opportunité', 'Opportunity Number', 'Date de création', 'Date de clôture', 'Étape', 'Revendeur', 'Client Final', 'Nom du produit', 'Code du produit', 'Quantité', 'Prix de vente', 'Prix total', 'Grossiste',
    # Added for manual change
    # 'Category Deal\nStock /CTO', 'Product Family\n(NX, NB, NR, PD, PT, PF)', 'Qty\nUnit', 'Revenu projet\nK Euros', 'Quarter Invoice\nFacturation', 'Forecast projet\nMenu déroulant', 'Next Step & Support demandé / Commentaire'
 
    # Cleanup OPTY and Model Name (remove NaN)
    df_master['Opportunity Number'].fillna("", inplace=True)
    df_master['Nom du produit'].fillna("", inplace=True)

    # Create Key Columns (Opty+Model)
    df_master['Key'] = df_master.apply(lambda row: f'{row["Opportunity Number"]}{row["Nom du produit"]}', axis = 1)

    # Copy df-master Columns value in df_pipe if exists already ... Otherwise leave blank
    # Column Category Deal
    df_pipe['Category Deal\nStock /CTO'] = df_pipe['Key'].map(Mapping_CatDeal)

    # Column Product Familly
    df_pipe['Product Family\n(NX, NB, NR, PD, PT, PF)'] = df_pipe['Key'].map(Mapping_ProdFam)

    # Column Quantity
    df_pipe['Qty\nUnit'] = df_pipe['Key'].map(Mapping_Qty)

    # Column Revenu projet
    df_pipe['Revenu projet\nK Euros'] = df_pipe['Key'].map(Mapping_RevEur)

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

    df_pipe.columns = df_master.columns

    worksheet.delete_rows(3, amount=(worksheet.max_row - 2))

#    wso = myworkbook.create_sheet("Updated Pipe")
    for r in dataframe_to_rows(df_pipe, index=False, header=False):
        worksheet.append(r)

    print(f'  - l onglet contient {len(df_master)} lignes maintenant')

    # Apply Columns Formats
    # Col C = 3
    Format_Cell(worksheet,3,numbers.FORMAT_DATE_DDMMYY)
    # Col C = 4
    Format_Cell(worksheet,4,numbers.FORMAT_DATE_DDMMYY)

    # Col K = 11
    Format_Cell(worksheet,11,numbers.FORMAT_CURRENCY_EUR_SIMPLE)
    # Col L = 12
    Format_Cell(worksheet,12,numbers.FORMAT_CURRENCY_EUR_SIMPLE)
    # Col Q = 17
    Format_Cell(worksheet,17,numbers.FORMAT_CURRENCY_EUR_SIMPLE)

    myworkbook.save(OUTPUT_SUIVI_RAW)

    print(f'- Sauvegarde vers {OUTPUT_SUIVI_RAW}')


if __name__ == "__main__":
    main()