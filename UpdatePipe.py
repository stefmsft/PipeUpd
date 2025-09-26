"""
PipeUpdUV - Salesforce Pipeline Data Integration Tool

This script integrates Salesforce pipeline export data (XLS/XLSX) into an existing Excel
tracking file (XLSM) for B2B sales opportunity management. The system focuses on
Opportunity (OpTY), Quotes, and Claims tracking while preserving manually entered data.

Version: 2.0

Enhancement Summary:
- Added comprehensive logging system with rotating file handler
- Implemented custom exception classes for better error categorization
- Enhanced data validation and sanitization functions
- Added configuration validation on startup
- Optimized data processing for better performance
- Added type hints throughout the codebase for maintainability
- Improved error handling with try-catch blocks
- Enhanced documentation with detailed docstrings
Author: PipeUpdUV Team
Last Modified: 2025

Key Features:
- Enhanced error handling and logging
- Data validation and sanitization
- Performance optimizations
- Configuration validation
- Type hints for better code maintainability

Usage:
    python UpdatePipe.py [file_path|all]

    Arguments:
        file_path: Process specific pipe file
        all: Process all files in the pipe directory
        (no args): Process latest pipe file
"""

import math
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import numbers
from datetime import date,datetime
import pandas as pd
import openpyxl
import glob
import os
import warnings
import sys
import time
import re
import shutil
import logging
from typing import List, Optional, Tuple, Dict, Any
from pathlib import Path
from dotenv import load_dotenv

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('updatepipe.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

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

NORMAXDELTA = os.getenv("NORMAXDELTA")
if (NORMAXDELTA == None):
    NORMAXDELTA=10000000
else:
    NORMAXDELTA = int(NORMAXDELTA)

ROLLINGWINDOWS = os.getenv("ROLLINGWINDOWS")
if (ROLLINGWINDOWS == None):
    ROLLINGWINDOWS=31
else:
    ROLLINGWINDOWS = int(ROLLINGWINDOWS)

ROLLINGFIELD = os.getenv("ROLLINGFIELD")
if (ROLLINGFIELD == None): ROLLINGFIELD='Date'

BCKUP_PIPE_FILE = (str(os.getenv("BCKUP_PIPE_FILE")).lower() == 'true' )
if (BCKUP_PIPE_FILE == None): BCKUP_PIPE_FILE=False

if BCKUP_PIPE_FILE:
    BCKUP_DIRECTORY = os.getenv("BCKUP_DIRECTORY")
    if (BCKUP_DIRECTORY == None):
        BCKUP_PIPE_FILE=False
    else:
        BCKUP_PIPE_FILE = True

if BCKUP_PIPE_FILE:
    BCKUP_GRANULARITY = os.getenv("BCKUP_GRANULARITY")
    if (BCKUP_GRANULARITY == None): BCKUP_GRANULARITY="Days"

CURWEEK = os.getenv("CURWEEK")
if (CURWEEK != None):
    CURWEEK = int(CURWEEK)
else:
    CURWEEK = None

# To avoid localisation colision
# Define col index for labels in Pipe file
# Only done for col name with problem
COL_OPTYOWNER=0
COL_CREATED=1
COL_CLOSED=2
COL_STAGE=3
COL_CUSTOMER=6
COL_TOTPRICE=9
COL_SALESMODELNAME=10

################################################################
# Exception Classes
################################################################
class PipeProcessingError(Exception):
    """Custom exception for pipe processing errors"""
    pass

class ConfigurationError(Exception):
    """Custom exception for configuration errors"""
    pass

class DataValidationError(Exception):
    """Custom exception for data validation errors"""
    pass

################################################################
# Configuration Validation
################################################################
def validate_configuration() -> None:
    """Validate all required configuration parameters"""
    required_configs = {
        'DIRECTORY_PIPE_RAW': DIRECTORY_PIPE_RAW,
        'INPUT_SUIVI_RAW': INPUT_SUIVI_RAW,
        'OUTPUT_SUIVI_RAW': OUTPUT_SUIVI_RAW
    }

    missing_configs = [key for key, value in required_configs.items() if not value]
    if missing_configs:
        raise ConfigurationError(f"Missing required configuration: {', '.join(missing_configs)}")

    # Validate paths exist
    if not os.path.exists(DIRECTORY_PIPE_RAW):
        raise ConfigurationError(f"Pipe directory does not exist: {DIRECTORY_PIPE_RAW}")

    if not os.path.exists(INPUT_SUIVI_RAW):
        raise ConfigurationError(f"Input tracking file does not exist: {INPUT_SUIVI_RAW}")

    # Validate numeric configurations
    if ROLLINGWINDOWS > 31:
        logger.warning(f"ROLLINGWINDOWS value {ROLLINGWINDOWS} exceeds recommended maximum of 31")

    logger.info("Configuration validation completed successfully")

################################################################
# Functions Helper
################################################################
def BackupPipeBefore(pipefullpath: str) -> bool:

    now = datetime.now()

    # dd/mm/YY
    dtstr = now.strftime("%Y%m%d-%H")

    filename = os.path.basename(pipefullpath).split('.')
    if len(filename) == 2:
        targetbckfn = f'{BCKUP_DIRECTORY}\\{filename[0]}-{dtstr}-bck.{filename[1]}'
    else:
        logger.error(f'Error building backup filename from {pipefullpath}')
        return False

    # Search on the exact name or on a part including only the day
    if BCKUP_GRANULARITY.lower() == "days":
        Targetsrch = f'{BCKUP_DIRECTORY}\\{filename[0]}-{now.strftime("%Y%m%d")}*-bck.{filename[1]}'
    else:
        Targetsrch = targetbckfn

    files = glob.glob(Targetsrch)

    if len(files) == 0:
        # Backup Pipe File
        try:
            shutil.copy(pipefullpath, targetbckfn)
            logger.info(f'Backup file created: {targetbckfn}')
        except Exception as e:
            logger.error(f'Failed to create backup: {str(e)}')
            return False

    return True

def GetLatestPipe(idir: str) -> str:
    """Get the latest pipe file from directory"""
    try:
        files = glob.glob(f'{idir}/*.xls*')
        if not files:
            raise PipeProcessingError(f"No Excel files found in directory: {idir}")

        latest_file = max(files, key=os.path.getctime)
        logger.info(f"Latest pipe file found: {latest_file}")
        return latest_file
    except Exception as e:
        logger.error(f"Error finding latest pipe file: {str(e)}")
        raise PipeProcessingError(f"Failed to find latest pipe file: {str(e)}")

def GetAllPipe(idir: str) -> List[str]:
    """Get all pipe files from directory, sorted by creation time"""
    try:
        files = glob.glob(f'{idir}/*.xls*')
        if not files:
            raise PipeProcessingError(f"No Excel files found in directory: {idir}")

        files.sort(key=os.path.getctime)
        logger.info(f"Found {len(files)} pipe files")
        return files
    except Exception as e:
        logger.error(f"Error getting all pipe files: {str(e)}")
        raise PipeProcessingError(f"Failed to get pipe files: {str(e)}")


def CheckPipeFile(pfile: str) -> bool:
    """Check if pipe file is valid Excel file"""
    try:
        if not os.path.isfile(pfile):
            logger.error(f"File does not exist: {pfile}")
            return False

        ext = os.path.splitext(pfile)[-1].lower()
        if ext not in ['.xls', '.xlsx']:
            logger.error(f"Invalid file extension: {ext}. Expected .xls or .xlsx")
            return False

        # Try to read the file to ensure it's not corrupted
        try:
            pd.read_excel(pfile, nrows=1)
        except Exception as e:
            logger.error(f"File appears to be corrupted: {str(e)}")
            return False

        logger.debug(f"Pipe file validation successful: {pfile}")
        return True
    except Exception as e:
        logger.error(f"Error validating pipe file {pfile}: {str(e)}")
        return False

################################################################
# Data Validation Functions
################################################################
def validate_dataframe_structure(df: pd.DataFrame, expected_cols: List[str], df_name: str) -> None:
    """Validate that dataframe has expected structure"""
    if df.empty:
        raise DataValidationError(f"{df_name} is empty")

    missing_cols = set(expected_cols) - set(df.columns)
    if missing_cols:
        logger.warning(f"{df_name} missing expected columns: {missing_cols}")

def sanitize_numeric_value(value: Any, default: float = 0.0) -> float:
    """Safely convert value to numeric, with fallback"""
    try:
        if pd.isna(value) or value == '':
            return default
        # Remove currency symbols and formatting
        if isinstance(value, str):
            cleaned = re.sub(r'[^\d.-]', '', str(value))
            return float(cleaned) if cleaned else default
        return float(value)
    except (ValueError, TypeError):
        logger.warning(f"Could not convert '{value}' to numeric, using default {default}")
        return default

def sanitize_date_value(value: Any) -> Optional[datetime]:
    """Safely convert value to datetime"""
    try:
        if pd.isna(value) or value == '':
            return None
        return pd.to_datetime(value, format='mixed')
    except Exception as e:
        logger.warning(f"Could not convert '{value}' to datetime: {str(e)}")
        return None

def sanitize_string_value(value: Any, default: str = '') -> str:
    """Safely convert value to string"""
    try:
        if pd.isna(value):
            return default
        return str(value).strip()
    except Exception:
        return default

#Mapping Date to Quarter FYear
def GetQFFromDate(cdate: datetime) -> Tuple[int, str]:

    cm = cdate.month
    if cm < 4:
        Quarter = 1
    else:
        if cm < 7:
            Quarter = 2
        else:
            if cm < 10:
                Quarter = 3
            else:
                Quarter = 4

    Year = str(cdate.year)[-2:]

    return Quarter, Year


#Generic Mapping Functions
def Mapping_Generic(Key: str, Col: str) -> str:
    """Generic mapping function to get value from master dataframe"""
    try:
        if df_master is None or df_master.empty:
            return ''

        rowval = df_master.loc[df_master['Key'] == Key]
        if len(rowval) == 0:
            return ''

        rtv = rowval.at[rowval.index[-1], Col]
        return sanitize_string_value(rtv)
    except Exception as e:
        logger.debug(f"Error in Mapping_Generic for Key {Key}, Col {Col}: {str(e)}")
        return ''

#Mapping Functions for
# 'Estimated\nQuantity', 'Revenu From\nEstinated Qty', 'Quarter Invoice\nFacturation', 'Forecast projet\nMenu déroulant', 'Next Step & Support demandé / Commentaire'

def Mapping_Qty(Key: str) -> Any:
    """Map quantity values with formula handling"""
    try:
        eq = Mapping_Generic(Key, 'Estimated\nQuantity')

        if str(eq).startswith('=') or str(eq) == '':
            rowval = df_master.loc[df_master['Key'] == Key]
            if len(rowval) > 0 and 'Quantité' in rowval.columns:
                eq = rowval['Quantité'].values[0]

        return sanitize_numeric_value(eq) if eq else ''
    except Exception as e:
        logger.debug(f"Error in Mapping_Qty for Key {Key}: {str(e)}")
        return ''

def Mapping_RevEur(Key: str) -> Any:
    """Map revenue values with formula and calculation handling"""
    try:
        rev = Mapping_Generic(Key, 'Revenu From\nEstinated Qty')

        if rev and rev != '':
            rowval = df_master.loc[df_master['Key'] == Key]
            if len(rowval) == 0:
                return ''

            if str(rev).startswith('='):
                if 'Prix total' in rowval.columns:
                    rev = sanitize_numeric_value(rowval['Prix total'].values[0])
            else:
                # Calculate from quantity and price
                if 'Estimated\nQuantity' in rowval.columns and 'Prix de vente' in rowval.columns:
                    qty = sanitize_numeric_value(rowval['Estimated\nQuantity'].values[0])
                    price = sanitize_numeric_value(rowval['Prix de vente'].values[0])
                    rev = qty * price

        return rev if rev else ''
    except Exception as e:
        logger.debug(f"Error in Mapping_RevEur for Key {Key}: {str(e)}")
        return ''

def Mapping_QtrInvoice(Key: str) -> str:
    """Map quarter invoice values with automatic calculation from close date

    Args:
        Key: Unique key for the opportunity

    Returns:
        Quarter invoice string in format QnFYyy or preserved value
    """

     # Rules :
     # If nothing, leave nothing
     # if lenght of value is not 2,4 or 6 put nothing
     # If first letter of value is Q, get the close date and calculate the QnFYyy

    eq = Mapping_Generic(Key,'Quarter Invoice\nFacturation')

    seq = str(eq)

    # Update : We translate the Close Date in QnFy even if the field is blank - Then it can eventually be changed. As long as it is in the  right format this wont be changed here.
    try:

        CloseDate = Mapping_Generic(Key,'Date de clôture')
        if str(CloseDate) != '':
            Quarter,Year = GetQFFromDate(CloseDate)
            seq = f'Q{Quarter}FY{Year}'

    except:
        pass


    return seq

def Mapping_FrCast(row: pd.Series) -> str:
    """Map forecast values based on Win Rate and Stage

    Args:
        row: Pandas Series containing opportunity data

    Returns:
        Forecast category string
    """

    Key = row['Key']

    eq = Mapping_Generic(Key,'Forecast projet\nMenu déroulant')
    seq = str(eq)

    AS = ["LOST = Perdu", "UNCOMMITED = Pas certain", "UNCOMMITED UPSIDE = Certain à 50% du WIN","COMMIT AT RISK = Certain à 75% du WIN","COMMIT = Certain à 100% du WIN","WIN = Gagné"]

# Update : Automatic fill of the column value base on Win Rate column ... If not empty
    fcast = seq

    Stat = row['Stage']

    if Stat.lower() == 'closed won':
        fcast = "WIN = Gagné"
    else:
        if seq not in AS:
            try:
                # rowval = df_master.loc[df_master['Key'] == Key]
                # WR = rowval['Win Rate'].values[0].replace('%', '')
                WR = row['Win Rate'].replace('%', '')
                # Check if the value is NaN or space (after converting to string and stripping whitespace)
                if pd.isna(WR) or str(WR).strip() == '':
                    return ''
                
                # Convert the value to a float to handle numeric comparison
                WR = float(WR)
                
                # Number of ranges is the same as the length of the AS array
                num_ranges = len(AS)
                range_size = 100 / num_ranges
                
                # Find the index in the AS array that corresponds to the value
                index = min(int((WR - 1) / range_size), num_ranges - 1)
                fcast = AS[index]

            except:
                pass

    return fcast

def Mapping_NxtStp(Key: str) -> str:
    """Map next step comments from existing data

    Args:
        Key: Unique key for the opportunity

    Returns:
        Next step comment string
    """

    return Mapping_Generic(Key,'Next Step & Support demandé / Commentaire')

def GetDynamicWeekColumns() -> List[str]:
    """Generate the 5 dynamic week column names based on current week

    Returns:
        List of 5 week column names: [Week-2, Week-1, Week, Week+1, Week+2]
    """
    if CURWEEK is not None:
        current_week = CURWEEK
        logger.debug(f'Using test week number: {current_week}')  # Only log at debug level to reduce noise
    else:
        current_week = datetime.now().isocalendar()[1]
        # Don't log here, let the calling code log the week columns being used

    week_columns = []

    for offset in range(-2, 3):  # -2, -1, 0, +1, +2
        week_num = current_week + offset
        # Handle year boundaries - most years have 52 weeks, some have 53
        if week_num < 1:
            # Get the actual number of weeks in the previous year
            prev_year = datetime.now().year - 1
            last_week_prev_year = datetime(prev_year, 12, 28).isocalendar()[1]  # Week containing Dec 28 is always the last week
            week_num = last_week_prev_year + week_num
        elif week_num > 52:
            # Check if current year actually has 53 weeks
            current_year = datetime.now().year
            last_week_current_year = datetime(current_year, 12, 28).isocalendar()[1]
            if week_num > last_week_current_year:
                week_num = week_num - last_week_current_year

        week_columns.append(f"Week {week_num}")

    return week_columns

def Mapping_WeekColumn(Key: str, old_col_name: str, new_col_name: str) -> str:
    """Preserve data from existing week columns when renaming

    Args:
        Key: Unique key for the opportunity
        old_col_name: Previous column name
        new_col_name: New column name

    Returns:
        Preserved data value or empty string
    """
    try:
        # First check if the new column name already exists in master
        if new_col_name in df_master.columns:
            return Mapping_Generic(Key, new_col_name)
        # If not, check if the old column name exists and map from it
        elif old_col_name in df_master.columns:
            return Mapping_Generic(Key, old_col_name)
        else:
            return ''  # New column, no existing data
    except:
        return ''

def Format_Cell(WS: openpyxl.worksheet.worksheet.Worksheet, start: int, ColIdx: int, Format: str) -> None:
    """Apply number format to a range of cells in a worksheet

    Args:
        WS: Worksheet to format
        start: Starting row number
        ColIdx: Column index to format
        Format: Number format string
    """
    try:
        for r in range(start, WS.max_row + 1):
            WS.cell(r, ColIdx).number_format = Format
        logger.debug(f"Applied format {Format} to column {ColIdx} starting from row {start}")
    except Exception as e:
        logger.error(f"Error formatting cells: {str(e)}")

def Write2Log(wb: openpyxl.Workbook, DataLst: List[Any]) -> pd.DataFrame:
    """Write pipeline data to the log sheet

    Args:
        wb: Excel workbook to write to
        DataLst: List containing [Date, Week, Nb OPTY, Sales Force Amount, Estimated Amount]

    Returns:
        DataFrame containing the updated log data
    """

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

    return df_log

def UpdatePipeAnalysis(wb: openpyxl.Workbook, df_log: pd.DataFrame) -> bool:
    """Update the Pipe Analysis sheet with rolling window analysis

    Args:
        wb: Excel workbook containing the analysis sheet
        df_log: DataFrame with log data for analysis

    Returns:
        Boolean indicating success/failure
    """
    # df_log expected columns:
    # 'Date', 'WK', 'Nb OPTY', 'Sales Force Amount', 'Estimated Amount'

    # Row where the Data starts (Generally 2 when the first row is used for header)
    LOGSHIFTROWDATA=2
    # Difference in row between the data start in "Log" tab versus the "analysis" tab
    # For instance 1 means that in the "Log" tab data starts row LOGSHIFTROWDATA=2 but on the
    # "Analysis" Tab it begins row 3
    SHIFTROWBETWEENTAB=1

    ret = False
    NormalizeEstimate = False

    shl = wb.sheetnames
    if "Pipe Log" not in shl:
        return ret

    # Get the Pipe Log Sheet
    wslog = wb['Pipe Log']

    if ROLLINGFIELD == 'WK':
        df_log = df_log.drop_duplicates(subset=['WK'], keep='last', ignore_index=False).copy()

    # Slicing for the last n values
    df_log = df_log.tail(ROLLINGWINDOWS).copy()
    #df_log.reset_index(inplace=True)

    # Get order of magnitude for the sales numbers
    # df_log['Magnitude'] = df_log.apply(lambda row: math.floor(math.log10(row['Sales Force Amount'])), axis = 1)
    MaxSFA = max(df_log['Sales Force Amount'])
    MinSFA = min(df_log['Sales Force Amount'])
    Mag = math.floor(math.log10(MaxSFA))

    # We substract according to its level of magnitude all common digit in the Amount serie
    # For instance if all 9 digits (Magnitude 8) amount start with 16 we substract 160000000 to the amount on the whole serie
    # The goal of the folowing loop is to find this subracted amount

    NormalizationVal = 0
    for d in range(Mag):
        df_log['Digit'] = df_log.apply(lambda row: str(row['Sales Force Amount'])[d], axis = 1)
        if len(df_log['Digit'].unique()) == 1:
            NormalizationVal = NormalizationVal + int(df_log['Digit'].unique()[0]) * 10**Mag
            Mag = Mag -1
        else:
            break

    # Check if a Normalization is also needed on the Estimated value
    # If min amount of the normalized value of Sales Force Amount is lower than the Estinate value we need to Normalize the the Estimate as well
    # For that we apply the same algorythme than before and verify that resulting difference between the 2 Normalized serie remain in the NORMAXDELTA range

    MaxSFAE = max(df_log['Estimated Amount'])
    MinSFAE = min(df_log['Estimated Amount'])
    Mag = math.floor(math.log10(MaxSFAE))

    NormalizationEVal = 0
    if MaxSFA  - NormalizationVal < MinSFAE:
        NormalizeEstimate = True
        for d in range(Mag):
            df_log['Digit'] = df_log.apply(lambda row: str(row['Estimated Amount'])[d], axis = 1)
            if len(df_log['Digit'].unique()) == 1:
                NormalizationEVal = NormalizationEVal + int(df_log['Digit'].unique()[0]) * 10**Mag
                Mag = Mag -1
            else:
                break
    # Check if Delta of Normalized value is to big (bigger than 4M)
    if (MinSFA - NormalizationVal) - (MaxSFAE - NormalizationEVal) > NORMAXDELTA:
        NormalizationVal = NormalizationVal + ((MinSFA - NormalizationVal) - (MinSFAE - NormalizationEVal) -  NORMAXDELTA)

    #Get the Pipe Log Sheet
    wsanalog = wb['Pipe Analysis']

    # Ecriture de la valeur de normalization
    # To make it more flexible la formule utilize une soustraction la valeur d'une celule fixe (R2, R=2,C=16)

    wsanalog.cell(row=2, column=18).value = NormalizationVal
    
    if NormalizeEstimate:
        wsanalog.cell(row=3, column=18).value = NormalizationEVal
    else:
        wsanalog.cell(row=3, column=18).value = 0

    # Write info on Run context
    wsanalog.cell(row=5, column=17).value = f'By {ROLLINGFIELD}'
    wsanalog.cell(row=6, column=17).value = f'For the last {ROLLINGWINDOWS} values'
    logger.info(f'Pipe Analysis with granularity on {ROLLINGFIELD}, showing the last {ROLLINGWINDOWS} records')
    # Write info on Run context
    wsanalog.cell(row=8, column=17).value = f'Last run : {date.today()}'

    wsanalog.cell(row=2, column=7).value = round(MaxSFA,0)
    wsanalog.cell(row=2, column=10).value = int(df_log['Sales Force Amount'].tail(1).iloc[0])

    #Effacement des valeurs precedentes
    for r in range(LOGSHIFTROWDATA,34):
        #Date Col = 1
        wsanalog.cell(row=(r+SHIFTROWBETWEENTAB), column=1).value = ''
        #Nb OPTY Col = 2
        wsanalog.cell(row=(r+SHIFTROWBETWEENTAB), column=2).value = ''
        #Opt Valorisation Col = 3
        wsanalog.cell(row=(r+SHIFTROWBETWEENTAB), column=3).value = ''
        #Opt Valorisation Col = 4
        wsanalog.cell(row=(r+SHIFTROWBETWEENTAB), column=4).value = ''
        #Ratio XForm Pipe Col = 16
        wsanalog.cell(row=(r+SHIFTROWBETWEENTAB), column=16).value = ''

    # Ecriture des formules dans les cellule sources du graph
    for i,r in enumerate(df_log.index):
        #Date Col = 1
        Formula = f"='Pipe Log'!A{r+1}"
        wsanalog.cell(row=(i+LOGSHIFTROWDATA+SHIFTROWBETWEENTAB), column=1).value = Formula
        #Nb OPTY Col = 2
        Formula = f"='Pipe Log'!C{r+1}"
        wsanalog.cell(row=(i+LOGSHIFTROWDATA+SHIFTROWBETWEENTAB), column=2).value = Formula
        #Opt Valorisation Col = 3
        Formula = f"='Pipe Log'!D{r+1}-$R$2"
        wsanalog.cell(row=(i+LOGSHIFTROWDATA+SHIFTROWBETWEENTAB), column=3).value = Formula
        #Opt Valorisation Col = 4
        Formula = f"='Pipe Log'!E{r+1}-$R$3"
        wsanalog.cell(row=(i+LOGSHIFTROWDATA+SHIFTROWBETWEENTAB), column=4).value = Formula
        #Ratio XForm Pipe Col = 16
        Formula = f"='Pipe Log'!E{r+1}/'Pipe Log'!D{r+1}"
        wsanalog.cell(row=(i+LOGSHIFTROWDATA+SHIFTROWBETWEENTAB), column=16).value = Formula

    # Cells Formating
    Format_Cell(wsanalog,3,1,numbers.FORMAT_DATE_DDMMYY)
    Format_Cell(wsanalog,3,3,'#,##0_-')
    Format_Cell(wsanalog,3,4,'#,##0_-')
    Format_Cell(wsanalog,3,16,'0%')
    wsanalog.cell(2,7).number_format = '[$EUR ]#,##0_-'
    wsanalog.cell(2,10).number_format = '[$EUR ]#,##0_-'

    return ret

def UpdatePipe(LatestPipe: str) -> None:
    """Main function to update pipe data with enhanced error handling"""
    global df_master

    logger.info(f"Starting pipe update process with file: {LatestPipe}")

    try:
        # Validate input file
        if not CheckPipeFile(LatestPipe):
            raise PipeProcessingError(f"Invalid pipe file: {LatestPipe}")

        # Row where the Data starts (Generally 2 when the first row is used for header)
        HEADERSHIFT=3

        # Get creation Date for futur usage in the Log Tab
        ctimef = datetime.strptime(time.ctime(os.path.getctime(LatestPipe)), "%a %b %d %H:%M:%S %Y")

        ####################################
        # Load Latest Pipe File
        ####################################

        logger.info(f'Using pipe file: {LatestPipe}')
        # Skip SKIP_ROW if extract made with header details. Depending on the header lines this value can be updated from .env file
        try:
            df_pipe = pd.read_excel(LatestPipe, skiprows=SKIP_ROW)
            logger.info(f"Successfully loaded pipe file with {len(df_pipe)} initial rows")
        except Exception as e:
            raise PipeProcessingError(f"Failed to read Excel file {LatestPipe}: {str(e)}")

        # Drop Empty Columns (more efficient with list comprehension)
        unnamed_cols = [col for col in df_pipe.columns if str(col).startswith('Unnamed:')]
        if unnamed_cols:
            df_pipe.drop(columns=unnamed_cols, inplace=True)
            logger.debug(f"Dropped {len(unnamed_cols)} unnamed columns")

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

        # Remove bogus values more efficiently with single operation
        bogus_values = ['Total', 'Confidential Information - Do Not Distribute',
                        'Copyright © 2000-2023 salesforce.com, inc. All rights reserved.']

        # Create mask for all bogus values at once
        mask = ~df_pipe[cols[COL_OPTYOWNER]].isin(bogus_values)
        df_pipe = df_pipe[mask]

        # Drop NaN values
        df_pipe.dropna(subset=[cols[COL_OPTYOWNER], cols[COL_CUSTOMER]], inplace=True)
        logger.debug(f"Removed bogus values and NaN entries")


        logger.info(f'Pipe file contains {len(df_pipe)} rows after initial processing')

    # Owner to keep
    # 'William ROMAN', 'Corinne CORDEIRO', 'Kajanan SHAN', 'Younes Giaccheri', 'Aziz ABELHAOU', 'Hippolyte FOVIAUX', 'Hatem ABBACI', 'Mehdi Dahbi', 'Gwenael BOJU', 'Charles TEZENAS', Etc ...
    
        # Owner filtering (more efficient with single isin operation)
        excluded_owners = ['Clement VIEILLEFONT', 'Vincent HALLER', 'Mathieu LUTZ', 'Calvin Chao']
        owner_mask = ~df_pipe[cols[COL_OPTYOWNER]].isin(excluded_owners)
        df_pipe = df_pipe[owner_mask]
        logger.debug(f"Filtered out excluded owners")

        # Client to Drop
        # 'Generic End User'
        # if Estimated Tot Price < 50K
        COL_SALESPRICE = 8
        COL_TOTPRICE = 9
        mask = (df_pipe[cols[COL_TOTPRICE]] < 50000) & df_pipe[cols[COL_CUSTOMER]].str.startswith('Generic')
        df_pipe.drop(df_pipe[mask].index, inplace=True)
        # Remove also the blank tot price for those 'Generic'
        mask = (df_pipe[cols[COL_TOTPRICE]]).isna() & df_pipe[cols[COL_CUSTOMER]].str.startswith('Generic')
        df_pipe.drop(df_pipe[mask].index, inplace=True)

        # Remove product lines more efficiently
        excluded_product_lines = ['LM', 'MS', 'MR']
        product_mask = ~df_pipe['Product Line'].isin(excluded_product_lines)
        df_pipe = df_pipe[product_mask]
        df_pipe['Product Line'] = df_pipe['Product Line'].fillna("")
        logger.debug(f"Filtered out excluded product lines")
        # df_pipe['Product Line'].fillna("", inplace=True) - Deprecated 3.12


        # Cleanup OPTY (remove NaN)
        df_pipe['Opportunity Number'] = df_pipe['Opportunity Number'].fillna("")
        df_pipe[cols[COL_SALESMODELNAME]] = df_pipe[cols[COL_SALESMODELNAME]].fillna("")
        # df_pipe['Opportunity Number'].fillna("", inplace=True) - Deprecated 3.12
        # df_pipe[cols[COL_SALESMODELNAME]].fillna("", inplace=True) - Deprecated 3.12

        # Format dates with error handling
        try:
            df_pipe[cols[COL_CREATED]] = pd.to_datetime(df_pipe[cols[COL_CREATED]], format='mixed', errors='coerce')
            df_pipe[cols[COL_CLOSED]] = pd.to_datetime(df_pipe[cols[COL_CLOSED]], format='mixed', errors='coerce')
            logger.debug("Date columns formatted successfully")
        except Exception as e:
            logger.warning(f"Date formatting issues: {str(e)}")

        # Copy "Run Rate" Type  Deals - But don't delete the line from the main Dataframe
        df_pipe_RR = df_pipe.loc[df_pipe['Deal Type']=='Run Rate Deal'].copy()

        # Copy "Closed Lost" Opportunities - And remove them from the main Dataframe later
        df_pipe_CL = df_pipe.loc[df_pipe[cols[COL_STAGE]]=='Closed Lost'].copy()

        # Create Key Columns (Opty+Model)
        df_pipe['Key'] = df_pipe.apply(lambda row: f'{row["Opportunity Number"]}{row[cols[COL_SALESMODELNAME]]}', axis = 1)

        logger.info(f'{len(df_pipe)} rows after data cleanup')

        ####################################
        # If Backup option is activated ... Then backup the actual Pipe file before processing.
        # Naming : name of INPUT_SUIVI_RAW "-yymmdd-hh-bck.xlsx"
        ####################################

        if BCKUP_PIPE_FILE:
            BackupPipeBefore(INPUT_SUIVI_RAW)

        ####################################
        # Load PipeLine Excel File and convert the 'Pipeline Sell Out' Tab to DataFrame
        ####################################

        try:
            myworkbook = openpyxl.load_workbook(INPUT_SUIVI_RAW, keep_vba=False)
            worksheet = myworkbook['Pipeline Sell Out']
        except Exception as e:
            raise PipeProcessingError(f"Failed to load tracking workbook {INPUT_SUIVI_RAW}: {str(e)}")

        ####################################
        # Creation/Update onglet Run Rate Pipe
        ####################################

        shl = myworkbook.sheetnames
        if "Pipeline Run Rate" in shl:
            worksheet_RR= myworkbook['Pipeline Run Rate']

            worksheet_RR.delete_rows(2, amount=(worksheet_RR.max_row+1))

            for r in dataframe_to_rows(df_pipe_RR, index=False, header=False):
                worksheet_RR.append(r)

        ####################################
        # Creation/Update onglet Closed Lost Pipe
        ####################################

        if "Pipeline Close Lost" in shl:
            worksheet_CL= myworkbook['Pipeline Close Lost']
            worksheet_CL.delete_rows(2, amount=(worksheet_CL.max_row+1))
        else:
            # Create new sheet if it doesn't exist
            worksheet_CL = myworkbook.create_sheet("Pipeline Close Lost")
            # Copy header from main sheet
            for col_num, cell in enumerate(myworkbook['Pipeline Sell Out'][1], 1):
                worksheet_CL.cell(row=1, column=col_num, value=cell.value)

        for r in dataframe_to_rows(df_pipe_CL, index=False, header=False):
            worksheet_CL.append(r)

        df_master = pd.DataFrame(worksheet.values)

        logger.info(f'Loading Pipeline Sell Out sheet from {INPUT_SUIVI_RAW}')


        # Drop first row
        df_master.drop(index=df_master.index[0], axis=0, inplace=True)

        # Set column name from new first row
        df_master.columns = df_master.iloc[0]
        # Reset the Index
        df_master = df_master.reset_index(drop=True)

        logger.info(f'Master file contains {len(df_master) - 1} rows')
        logger.info('Starting opportunity injection/refresh...')

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
        df_master['Opportunity Number'] = df_master['Opportunity Number'].fillna("")
        df_master['Nom du produit'] = df_master['Nom du produit'].fillna("")

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
        # df_pipe['Forecast projet\nMenu déroulant'] = df_pipe['Key'].map(Mapping_FrCast)
        df_pipe['Forecast projet\nMenu déroulant'] = df_pipe.apply(Mapping_FrCast, axis=1)

        # Column Next Step
        df_pipe['Next Step & Support demandé / Commentaire'] = df_pipe['Key'].map(Mapping_NxtStp)

        # Dynamic Week Columns (5 columns: Week-2, Week-1, Week, Week+1, Week+2)
        dynamic_week_columns = GetDynamicWeekColumns()
        current_week = datetime.now().isocalendar()[1] if CURWEEK is None else CURWEEK
        logger.info(f'Adding dynamic week columns (current week {current_week}): {dynamic_week_columns}')

        # Get existing column names that might contain week data (for preservation)
        existing_week_columns = [col for col in df_master.columns if col and str(col).startswith('Week ')]

        for i, new_week_col in enumerate(dynamic_week_columns):
            # Try to preserve data from existing week columns if they exist
            if i < len(existing_week_columns):
                old_week_col = existing_week_columns[i]
                df_pipe[new_week_col] = df_pipe['Key'].apply(
                    lambda key: Mapping_WeekColumn(key, old_week_col, new_week_col)
                )
            else:
                # New column, initialize with empty values
                df_pipe[new_week_col] = df_pipe['Key'].map(
                    lambda key: Mapping_Generic(key, new_week_col)
                )

        # Remove "Étape:Rejected"
        df_pipe.drop(df_pipe.loc[df_pipe[cols[COL_STAGE]]=='Rejected'].index, inplace=True)

        # Remove "Closed Lost" opportunities (they are now in their own tab)
        df_pipe.drop(df_pipe.loc[df_pipe[cols[COL_STAGE]]=='Closed Lost'].index, inplace=True)

        # No need of the Key Column anymore
        df_pipe.drop(['Key'], axis=1, inplace=True)
        df_master.drop(['Key'], axis=1, inplace=True)

        SFPipeAmmount = df_pipe[cols[COL_TOTPRICE]].sum()
        df_pipe['Revenu-Val'] = df_pipe['Estimated\nQuantity'].apply(pd.to_numeric, errors='coerce')
        df_pipe['Revenu-Val'] = df_pipe['Revenu-Val'] * df_pipe[cols[COL_SALESPRICE]]
        EstPipeAmmount = df_pipe['Revenu-Val'].sum()
        df_pipe.drop('Revenu-Val',axis=1,inplace=True)

        # Clean up None columns more efficiently
        try:
            none_columns = [col for col in df_master.columns if col is None]
            if none_columns:
                df_master.drop(columns=none_columns, inplace=True)
                logger.debug(f"Removed {len(none_columns)} None columns from master")
        except Exception as e:
            logger.debug(f"Error cleaning None columns: {str(e)}")

        df_pipe.columns = df_master.columns

        worksheet.delete_rows(3, amount=(worksheet.max_row - 2))

        for r in dataframe_to_rows(df_pipe, index=False, header=False):
            worksheet.append(r)

        # Update Excel column headers for dynamic Week columns (starting at column V = 22)
        # Reuse the dynamic_week_columns already calculated above
        logger.info(f'Updating Excel column headers for Week columns: {dynamic_week_columns}')
        for i, week_col_name in enumerate(dynamic_week_columns):
            col_idx = 22 + i  # V=22, W=23, X=24, Y=25, Z=26
            worksheet.cell(row=2, column=col_idx).value = week_col_name

        for i in range(HEADERSHIFT,worksheet.max_row+1):
            worksheet.cell(i,18).value = f'=Q{i}*I{i}'

        logger.info(f'Updated sheet now contains {len(df_pipe)} rows')

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
        Format_Cell(worksheet,3,18,'[$EUR ]#,##0_-')

        # Log Pipe Data
        lst = [datetime(ctimef.year,ctimef.month,ctimef.day,0,0), ctimef.isocalendar()[1], worksheet.max_row - 2, SFPipeAmmount, EstPipeAmmount]
        logger.info(f'Updating Pipe Log with: {lst}')
        df_log = Write2Log(myworkbook,lst)

        if "Pipe Analysis" in myworkbook.sheetnames:
            logger.info('Refreshing Pipe Analysis sheet')
            UpdatePipeAnalysis(myworkbook,df_log)

        myworkbook.save(OUTPUT_SUIVI_RAW)
        logger.info(f'Saving to: {OUTPUT_SUIVI_RAW}')

    except Exception as e:
        logger.error(f"Error during pipe update: {str(e)}")
        raise PipeProcessingError(f"Pipe update failed: {str(e)}")

    logger.info("Pipe update completed successfully")
    return

def main() -> None:
    """Main function with comprehensive error handling and validation"""
    try:
        logger.info("Starting UpdatePipe application")

        # Validate configuration first
        validate_configuration()

        loopProc = False
        PipeFList = []

        if len(sys.argv) > 1:
            logger.info(f'Parameter {sys.argv[1]} detected')
            if sys.argv[1].lower() == 'all':
                loopProc = True
                PipeFList = GetAllPipe(DIRECTORY_PIPE_RAW)
                logger.info(f"Processing all {len(PipeFList)} pipe files")
            else:
                if CheckPipeFile(sys.argv[1]):
                    LatestPipe = sys.argv[1]
                    logger.info(f"Processing specific file: {LatestPipe}")
                else:
                    raise PipeProcessingError(f'Invalid Pipe file: {sys.argv[1]}')
        else:
            LatestPipe = GetLatestPipe(DIRECTORY_PIPE_RAW)
            logger.info(f"Processing latest file: {LatestPipe}")

        # Process files
        if loopProc:
            for i, f in enumerate(PipeFList, 1):
                logger.info(f"Processing file {i}/{len(PipeFList)}: {f}")
                UpdatePipe(f)
        else:
            UpdatePipe(LatestPipe)

        logger.info("UpdatePipe application completed successfully")

    except ConfigurationError as e:
        logger.error(f"Configuration error: {str(e)}")
        sys.exit(1)
    except PipeProcessingError as e:
        logger.error(f"Processing error: {str(e)}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()