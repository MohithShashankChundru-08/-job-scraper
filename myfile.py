import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

def get_job(URL):
    scraped_data = []
    try:
        response = requests.get(URL)
        response.raise_for_status()
    except Exception as e:
        print('Error in fetching website details:',e)
        return []
    soup = BeautifulSoup(response.text,'html.parser')
    for tag in soup.find_all('div',class_='card-content'):
        job_title = tag.find('h2',class_='title').text
        company = tag.find('h3',class_ = 'subtitle').text
        location = tag.find('p', class_='location').text.strip()
        date_posted = tag.find('time').text
        scraped_data.append({'Job Role':job_title,'Company':company,'Location':location,'Date Posted':date_posted})
    return scraped_data

def clean_data(scraped_data):
    df = pd.DataFrame(scraped_data)
    df = df.drop_duplicates()
    df = df.dropna(subset=['Job Role'])
    df['Company'] = df['Company'].fillna('NA')
    df['Location'] = df['Location'].fillna('NA')
    df['Date Posted'] = pd.to_datetime(df['Date Posted'], errors='coerce')
    df = df.dropna(subset=['Date Posted'])
    return df

def filter_data(df,keyword):
    if not df.empty:
        filtered = df[df['Job Role'].str.contains(keyword,case=False,na=False)].copy()
    else:
        filtered = df.copy()
    return filtered

def transform_data(filtered):
    filtered['Is Remote'] = filtered['Location'].apply(lambda x: 'Yes' if 'remote' in str(x).lower() else 'No')
    transformed = filtered.copy()
    return transformed
        
def sort_data(transformed):
    sorted_data = transformed.sort_values(by='Date Posted',ascending = False).copy()
    return sorted_data

def save_data(sorted_data, filename):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = 'Jobs Details'

    ws1.append(list(sorted_data.columns))

    for row in sorted_data.itertuples(index=False):
        ws1.append(list(row))

    # Colors
    header_fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')  # Dark blue
    alt_fill = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')    # Light blue

    header_font = Font(color='FFFFFF', bold=True)
    normal_font = Font(color='000000')

    alignment = Alignment(horizontal='left', vertical='center')

    border_side = Side(style='thin', color='000000')
    border = Border(top=border_side, bottom=border_side, left=border_side, right=border_side)

    # Style header
    for col in range(1, len(sorted_data.columns) + 1):
        cell = ws1.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = alignment
        cell.border = border

    # Style rows (alternate colors)
    for row in range(2, len(sorted_data) + 2):
        for col in range(1, len(sorted_data.columns) + 1):
            cell = ws1.cell(row=row, column=col)

            if row % 2 == 0:
                cell.fill = alt_fill

            cell.font = normal_font
            cell.alignment = alignment
            cell.border = border

    ws2 = wb.create_sheet(title='Summary')
    remote_jobs = len(sorted_data[sorted_data['Is Remote']=='Yes'])
    not_remote_jobs = len(sorted_data[sorted_data['Is Remote']=='No'])
    ws2['A1'] = 'Metric'
    ws2['B1'] = 'Value'
    ws2['A2'] = "Total Jobs"
    ws2['B2'] = len(sorted_data)
    ws2['A3'] = 'Total Remote Jobs'
    ws2['B3'] = remote_jobs
    ws2['A4'] = 'Total Non-Remote Jobs'
    ws2['B4'] = not_remote_jobs
    cell_font = Font(name='Calibri',size=14,bold=False,color='000000')
    cell_alignment = Alignment(horizontal='left',vertical='top')
    cell_side = Side(style='thin',color='000000')
    cell_border = Border(top=cell_side,bottom=cell_side,left=cell_side,right=cell_side)
    for colu in range(1,len(sorted_data.columns) + 1):
        cell = ws1.cell(row=1, column=colu)
        cell.font = Font(name='Calibri',size=14,bold=True,color='FFFFFF')
        cell.fill = header_fill
        cell.alignment = cell_alignment
        cell.border = cell_border
    for ro in range(2, len(sorted_data) + 2):
        for co in range(1, len(sorted_data.columns) + 1):
            cell = ws1.cell(row=ro,column=co)
            cell.font = cell_font
            cell.alignment = cell_alignment
            cell.border = cell_border
    ws2['B1'].font = cell_font
    for col in range(1,5):
        ws2[f'A{col}'].font = Font(bold=True)
        ws2[f'A{col}'].alignment = cell_alignment
        ws2[f'B{col}'].alignment = cell_alignment
    ws3 = wb.create_sheet(title='Remote Jobs')
    remote_df = sorted_data[sorted_data['Is Remote'] == 'Yes']
    for row in remote_df.itertuples(index=False):
        ws3.append(list(row))
    ws3.append(list(remote_df.columns))
    ws3['A1'].font = Font(name='Calibri',size=14,bold=True)
    ws3['B1'].font = Font(name='Calibri',size=14,bold=True)
    ws3['C1'].font = Font(name='Calibri',size=14,bold=True)
    ws3['D1'].font = Font(name='Calibri',size=14,bold=True)
    ws3['E1'].font = Font(name='Calibri',size=14,bold=True)
    ws3['A1'].alignment = cell_alignment
    ws3['B1'].alignment = cell_alignment
    ws3['C1'].alignment = cell_alignment
    ws3['D1'].alignment = cell_alignment
    ws3['E1'].alignment = cell_alignment
    total_jobs = len(sorted_data)
    total_remote_jobs = len(sorted_data[sorted_data['Is Remote'] == 'Yes'])
    remote_percentage = (total_remote_jobs/total_jobs)*100 if total_jobs > 0 else 0
    ws2['A5'] = 'Total Remote Jobs'
    ws2['B5'] = round(remote_percentage,4)
    ws2['A5'].font = Font(name='Calibri',size=14,bold=True)
    ws2['A5'].alignment = cell_alignment
    ws2['B5'].alignment = cell_alignment
    top_companies = sorted_data['Company'].value_counts().head(3)
    ws2['A7'] = 'Top Companies'
    ws2['A8'] = 'Company'
    ws2['B8'] = 'Job Count'
    ws2['A7'].font = Font(name='Calibri',size=14,bold=True)
    ws2['A8'].font = Font(name='Calibri',size=14,bold=True)
    ws2['B8'].font = Font(name='Calibri',size=14,bold=True)
    ws2['A7'].alignment = cell_alignment
    ws2['A8'].alignment = cell_alignment
    ws2['B8'].alignment = cell_alignment
    row = 9
    for company, count in top_companies.items():
        ws2[f'A{row}'] = company
        ws2[f'B{row}'] = count
        ws2[f'A{row}'].alignment = cell_alignment
        ws2[f'B{row}'].alignment = cell_alignment
        row += 1
    ws4 = wb.create_sheet(title='Company Analysis')
    ws4.append(['Company', 'Total Jobs'])
    company_counts = sorted_data['Company'].value_counts()
    for company, count in company_counts.items():
        ws4.append([company, count])
    for row in range(2, len(set(sorted_data['Company'])) + 2):
        for col in range(1, 3):
            cell = ws4.cell(row=row, column=col)
            if row % 2 == 0:
                cell.fill = alt_fill
            cell.font = normal_font
            cell.alignment = alignment
            cell.border = border
    for colu in range(1,3):
        cell = ws4.cell(row=1, column=colu)
        cell.font = Font(name='Calibri',size=14,bold=True,color='FFFFFF')
        cell.fill = header_fill
        cell.alignment = cell_alignment
        cell.border = cell_border
    for ws in wb.worksheets:
        for col in ws.columns:
            max_length = 0
            column_letter = col[0].column_letter 
            for cell in col:
                try:
                    if cell.value:
                        length = len(str(cell.value))
                        if length > max_length:
                            max_length = length
                except:
                    pass
            adjusted_width = (max_length + 2)*1.2
            ws.column_dimensions[column_letter].width = adjusted_width
    wb.save(filename)
    print("Excel file created and saved successfully.")

def main():
    URL = 'https://realpython.github.io/fake-jobs/'
    filename = input('Enter your filename (eg. data.xlsx): ')
    jobs = get_job(URL)
    cleaned = clean_data(jobs)
    transformed = transform_data(cleaned)
    save_data(transformed,filename)

if __name__=='__main__':
    main()
