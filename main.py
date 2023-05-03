import pandas as pd
import datetime
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
import io
import json
import requests


#Gets a report from Quickbase API, parses the response, and returns a Pandas DF
def GetReport():
    headers = {
        'QB-Realm-Hostname': {yourcompany.quickbase.com},#Use your company's quickbase url
        'User-Agent': '{User-Agent}',
        'Authorization': 'QB-USER-TOKEN xxxxxx_xxxx_x_xxxxxxxxxxxxxxxxxxxxxxxxxx' #Replace with your Quickbase API User token
    }
    params = {
        'tableId': 'xxxxxxxxx', #Get the table id for the table you want (this can be found in the url when you are viewing the table in quickbase fpr example https://mycompany.quickbase.com/db/tableidhere?a=q&qid=reportidhere
       
    }
    r = requests.post(
    'https://api.quickbase.com/v1/reports/{report_number}/run', #Use the report number from your url like in the example above
    params = params, 
    headers = headers
    )
   
    report = json.dumps(r.json())
    
    data = json.loads(report)
    df = pd.json_normalize(data['data'])
    col_names = list(df.columns)
 
    new_col_names = {}
    for field in data['fields']:
        col_id = (str(field['id']) + '.value').strip()
        col_label = field['label']
        col_name = f"{col_id}"
       
        if col_name in col_names:
            new_col_names[col_name] = col_label
            
        
    df = df.rename(columns=new_col_names)
   
    return df



def writeFiles():
    #calls the GetReport function
    df = GetReport()

    #group the data together by desired column 
    df = df.sort_values('{desired grouping column}')

    groups = df['{desired grouping column'].unique()
    df = df.fillna('')
    grouped_dfs = {}

    for item in groups:
        group_df = df[df['{desired grouping column}'] == item]
        grouped_dfs[item] = group_df

    today = datetime.date.today()
    today = today.strftime('%m-%d-%y')
    for group_name, group_df in grouped_dfs.items():
        
        #Excel
        file_name = 'output/' + today + group_name.replace(' ', '') + '.xlsx'
        writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
        group_df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        worksheet.set_column('A:Z', 20)
        writer._save()

        #PDF
        file_name = 'output/' +group_name + '.pdf'
        pdf_buffer = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buffer, pagesize=landscape(letter), topMargin=0, bottomMargin=0)
        c = canvas.Canvas(pdf_buffer, pagesize=letter)
        elements = []
        styles = getSampleStyleSheet()
        table_data = [group_df.columns.tolist()] + group_df.values.tolist()
        t = Table(table_data, splitByRow=True)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D4D4D4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 4),
            ('BOTTOMPADDING', (0, 0), (0, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor('#000000')),
            ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (0, 0), 8),
        ]))
        elements.append(t)

        # Build the PDF document and save it to a file
        doc.build(elements)
        pdf_data = pdf_buffer.getvalue()

        # Write the PDF data to a file
        with open(file_name, 'wb') as f:
            f.write(pdf_data)

writeFiles()






