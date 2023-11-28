import pandas as pd
import numpy as np
from sql_engine import connect

df_2023 = pd.read_excel(r'C:\Medline\8. database\Asia Inspection Database\2022\QP-00017-F-00005 Asia Inspection Database 2023.XLSM',sheet_name="Sheet1")

vendor_mapping = pd.read_excel(r'C:\Medline\2. CPM\data\vendor_mapping\Vendor _mapping 2023_v1.xlsx')

manager_dict = dict(zip(vendor_mapping['Vendor Number'],vendor_mapping['Regional Manager']))
supervisor_dict = dict(zip(vendor_mapping['Vendor Number'],vendor_mapping['Supervisor']))
scheduler_dict = dict(zip(vendor_mapping['Vendor Number'],vendor_mapping['Scheduler']))

def assign_inspection_result(df):
    return df.assign(**{'Inspection Result' : lambda d : d['Results'].map({'A':'Inspected', 'R':'Inspected', 'W':'Waived', 'G':'Guaranteed'})})

std = 10
(
    df_2023.query('`Inspection Date`.dt.month == @std')
    .pipe(assign_inspection_result)
    .assign(**{
        "Supervisor" : lambda d : d['Vendor Code'].get(supervisor_dict),
        "Regional Manager" : lambda d : d['Vendor Code'].get(manager_dict),
        "Scheduler" : lambda d : d['Vendor Code'].get(scheduler_dict)
    })
)

