from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import openpyxl
import os
import sys

def Dedup(PipeF):

    df_pipe = pd.read_excel(PipeF)
    print (f'Taille avant Dedup {len(df_pipe)}')
    df_pipe.drop_duplicates(subset=['Opportunity Number', 'Prix total'], inplace=True)
    print (f'Taille apres Dedup {len(df_pipe)}')
    df_pipe.to_excel('Dedup.xlsx', index=False)

def main():

    PipeFileN = "C:\Projects\PipeUpd\Classeur1.xlsx"

    Dedup(PipeFileN)

    return

if __name__ == "__main__":
    main()