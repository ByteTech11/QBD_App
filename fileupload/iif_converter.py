import pandas as pd
from openpyxl import load_workbook
from datetime import datetime, timedelta
import re

def convert_excel_to_iif(excel_file, output_iif_path):
    try:
        df = pd.read_excel(excel_file, engine='openpyxl', skiprows=12)
    except Exception as e:
        return f"Error reading the Excel file: {str(e)}"

    df.columns = df.columns.str.strip() 

    if 'STHOURS' not in df.columns:
        df['STHOURS'] = df.get('Stat Hours', pd.Series([0] * len(df)))  

    try:
        wb = load_workbook(excel_file)
        sheet = wb.active
        date_from_excel = None

        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(row=4, column=col).value
            if isinstance(cell_value, str):
                match = re.findall(r'(\d{2}-[A-Za-z]{3}-\d{2})', cell_value)  
                if match and len(match) >= 2:
                    date_from_excel = match[1]  
                    break  
        
        if not date_from_excel:
            raise ValueError("No valid second date found in row 4.")

        extracted_date = datetime.strptime(date_from_excel, '%d-%b-%y')  
        
        adjusted_date = extracted_date - timedelta(days=2)

        formatted_date = adjusted_date.strftime('%m/%d/%Y') 
    except Exception as e:
        return f"Error extracting or adjusting the date: {str(e)}"

    df_cleaned = df[['Employee', 'Emp\nNum', 'Emp Type', 'Reg Hours', 'Stat Hours', 'PH Hours', 'Total', 'Rate', 'STHOURS']]
    df_cleaned = df_cleaned.dropna(subset=['Employee', 'Emp\nNum', 'Reg Hours', 'Total'])

    iif_header = "!TIMEACT\tDATE\tJOB\tEMP\tITEM\tPITEM\tDURATION\tPROJ\tNOTE\tXFERTOPAYROLL\tBILLINGSTATUS\n"
    iif_data = [iif_header]

    for _, row in df_cleaned.iterrows():
        timeact = "TIMEACT"
        date = formatted_date    
        job = "Default Job"
        emp = row['Employee']
        emp_num = row['Emp\nNum'] 
        emp_type = row['Emp Type'] if pd.notna(row['Emp Type']) else 0 
        item = "ST Rate"  
        pitem = "Regular Pay"
        duration = row['Total'] if pd.notna(row['Total']) else row['Reg Hours']  
        stat_hours = row['STHOURS'] if pd.notna(row['STHOURS']) else 0  
        ph_hours = row['PH Hours'] if pd.notna(row['PH Hours']) else 0 
        proj = ""
        note = ""  
        xfertopayroll = "Y"
        billingstatus = "1"

        if emp_type == 'IS': 
            if duration > 48:
                duration = 48  
        elif emp_type == 'SD': 
            if duration > 88:
                duration = 88  
        else:
            if duration > 88:
                duration = 88

        if duration > 0:
            iif_row_regular = f"{timeact}\t{date}\t{job}\t{emp}\t{item}\t{pitem}\t{duration}\t{proj}\t{note}\t{xfertopayroll}\t{billingstatus}"
            iif_data.append(iif_row_regular)

        if ph_hours > 0:
            item = "PH Hours"  
            pitem = "Public Holidays Hours"
            iif_row_ph = f"{timeact}\t{date}\t{job}\t{emp}\t{item}\t{pitem}\t{ph_hours}\t{proj}\t{note}\t{xfertopayroll}\t{billingstatus}"
            iif_data.append(iif_row_ph)

        if stat_hours > 0:
            item = "STAT Hours"  
            pitem = "Statutory Pay"
            iif_row_stat = f"{timeact}\t{date}\t{job}\t{emp}\t{item}\t{pitem}\t{stat_hours}\t{proj}\t{note}\t{xfertopayroll}\t{billingstatus}"
            iif_data.append(iif_row_stat)

    iif_content = "\n".join(iif_data)

    try:
        with open(output_iif_path, 'w') as iif_file:
            iif_file.write(iif_content)
        print(f"IIF file successfully created: {output_iif_path}")
    except Exception as e:
        return f"Error saving the IIF file: {str(e)}"

    return iif_content


excel_file_path = '/mnt/data/Your_Excel_File.xlsx'  
output_iif_path = '/mnt/data/output_file.iif' 

convert_excel_to_iif(excel_file_path, output_iif_path)
