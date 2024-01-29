import os
import openpyxl
import datetime
import pandas as pd
import win32com.client as win32
import warnings
warnings.filterwarnings("ignore")

def extract_formulas(file_path):
    wb = openpyxl.load_workbook(file_path, read_only=True, keep_vba=True)

    formulas_with_details = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and str(cell.value).startswith('='):
                    formula = str(cell.value)                    
                    coordinates = cell.coordinate
                    formulas_with_details.append((file_path, os.path.basename(file_path), sheet, coordinates, formula.replace('=','')))
    df = pd.DataFrame(formulas_with_details, columns=['Path', 'Excel Name', 'Sheet Name', 'Coordinates', 'Formulas'])
    return df

def extract_format(file_path):
    wb = openpyxl.load_workbook(file_path, read_only=True, keep_vba=True)

    formats_with_details = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and str(cell.value).startswith('='):
                    
                    formats_with_details.append((file_path, os.path.basename(file_path), sheet, coordinates, formula.replace('=','')))
    df = pd.DataFrame(formats_with_details, columns=['Path', 'Excel Name', 'Sheet Name', 'Coordinates', 'Formulas'])
    return df

countofflase = 0

def compare_formulas(df1, df2):
    global countofflase
    merged_df = pd.merge(df1, df2, on=['Sheet Name', 'Coordinates'], how='outer', suffixes=('_1', '_2'))
    merged_df['Status'] = merged_df['Formulas_1'] == merged_df['Formulas_2']
    countofflase = len(merged_df[merged_df['Status'] == False])
    return merged_df

def sendmail(recipients, file_path3, reportname=None, Location1=None, Location2=None, cc_recipients=None):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipients
    mail.CC = cc_recipients
    status = 'Success' if countofflase == 0 else 'Failure'
    mail.Subject = status + ' | ' + 'Excel Formula Comparison' + ' | ' + reportname
    attachment3 = file_path3
    mail.Attachments.Add(attachment3)
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    html_body = f'''
    <html>
    <body>
        <h2>Excel Formula Comparison Report</h2>
        <p><strong>Report Name:</strong> {reportname}</p>
        <p><strong>Location 1:</strong> {Location1}</p>
        <p><strong>Location 2:</strong> {Location2}</p>
        <p><strong>Status:</strong> <span style="color: {'green' if status == 'Success' else 'red'}"><strong>{status}</strong></span></p>
        <p><strong>No of not matching formulas:</strong> {countofflase}</p>
        <p><strong>Executed time:</strong> {current_time}</p>
        <br>
        <p>Thanks and Regards,</p>
        <p><i>Reporting Dev Team</i></p>
    </body>
    </html>
    '''
    mail.HTMLBody = html_body
    mail.Send()
    print("Mail sent successfully")

def formatting():
    pass


if __name__ == '__main__':
    file_path1 = 'C:\\Users\\v-mbnvsai\\OneDrive - Microsoft\\Desktop\\Dev\\test11.xlsm'
    file_path2 = 'C:\\Users\\v-mbnvsai\\OneDrive - Microsoft\\Desktop\\Prod\\test21.xlsm'
    reportname = os.path.basename(file_path1)
    Location1 = os.path.basename(os.path.dirname(file_path1))
    Location2 = os.path.basename(os.path.dirname(file_path2))
    df1 = extract_formulas(file_path1)
    df2 = extract_formulas(file_path2)
    merged_df = compare_formulas(df1, df2)
    merged_df.to_excel('formulas_comparison.xlsx', index=False)
    file_path3 = 'C:\\Users\\v-mbnvsai\\OneDrive - Microsoft\\Documents\\Learning\\Py\\Formulas\\formulas_comparison.xlsx'
    # recipients = 'saimanojb@maqsoftware.com'
    # # cc_recipients = 'saipavanm@maqsoftware.com'
    # # recipients = input("Enter recipients: ")
    # cc_recipients = input("Enter cc recipients: ")
    # sendmail(recipients, file_path3, reportname,Location1,Location2,cc_recipients)