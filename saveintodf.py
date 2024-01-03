import os
import openpyxl
import pandas as pd
import win32com.client as win32

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

countofflase = 0

def compare_formulas(df1, df2):
    merged_df = pd.merge(df1, df2, on=['Sheet Name', 'Coordinates'], how='outer', suffixes=('_1', '_2'))
    merged_df['Status'] = merged_df['Formulas_1'] == merged_df['Formulas_2']
    if(merged_df['Status'].all() == False):
        countofflase = countofflase + 1
    return merged_df

def sendmail(recipients, cc_recipients=None):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipients
    mail.CC = cc_recipients
    mail.Subject = 'Mail regarding validation'
    mail.Body = 'Hi Team, \n\nThe results for the validation of the given files is as follows\nThe formulas in the report ' + ('are same' if countofflase == 0 else 'are not same and count of formulas which are not same is ' + str(countofflase))
    # mail.HTMLBody = '<div><img src=""C:\\Users\\v-mbnvsai\\OneDrive - Microsoft\\Documents\\Learning\\Py\\Formulas\\1.png"" alt="Business Card Image"></div>' #this field is optional
    # mail.GetInspector
    # signature = mail.HTMLBody
    # mail.HTMLBody = signature
    mail.Send()
    print("Mail sent successfully")



if __name__ == '__main__':
    file_path1 = 'C:\\Users\\v-mbnvsai\\OneDrive - Microsoft\\Desktop\\test1.xlsm'
    file_path2 = 'C:\\Users\\v-mbnvsai\\OneDrive - Microsoft\\Desktop\\test2.xlsm'
    df1 = extract_formulas(file_path1)
    df2 = extract_formulas(file_path2)
    merged_df = compare_formulas(df1, df2)
    merged_df.to_excel('formulas_comparison.xlsx', index=False)
    recipients = 'saimanojb@maqsoftware.com'
    recipients = input("Enter recipients (separated by commas): ")
    cc_recipients = input("Enter cc recipients (separated by commas): ")
    sendmail(recipients, cc_recipients)


