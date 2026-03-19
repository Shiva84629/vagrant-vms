from django.shortcuts import render
import openpyxl
from datetime import datetime, date

def display_excel(request):
    # Path to the Excel file
    excel_path = r"D:\study material\dhqp\LotInfo.xlsx"
    
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(excel_path, data_only=True)  # data_only to get values, not formulas
        sheet = wb.active
        
        # Read all data
        all_data = []
        for row in sheet.iter_rows(values_only=True):
            all_data.append(row)
        
        # Skip header row if exists
        data_rows = all_data[1:] if len(all_data) > 1 else all_data
        
        if request.method == 'POST':
            search_date = request.POST.get('search_date')
            search_job = request.POST.get('search_job')
            show_all = request.POST.get('show_all')
            if show_all:
                filtered_data = data_rows
                search_term = "All Data"
            elif search_date or search_job:
                # Build filter conditions
                def matches_row(row):
                    date_match = True
                    job_match = True
                    
                    if search_date:
                        try:
                            search_dt = datetime.fromisoformat(search_date).date()
                            date_formats = [
                                search_dt.strftime('%Y-%m-%d'),
                                search_dt.strftime('%m/%d/%Y'),
                                search_dt.strftime('%d/%m/%Y'),
                                search_dt.strftime('%Y/%m/%d'),
                                search_dt.strftime('%d-%m-%Y'),
                                str(search_dt),
                            ]
                            date_match = any(
                                (isinstance(cell, (date, datetime)) and (
                                    (isinstance(cell, datetime) and cell.date() == search_dt) or
                                    (isinstance(cell, date) and cell == search_dt)
                                )) or
                                any(fmt in str(cell) for fmt in date_formats if cell)
                                for cell in row
                            )
                        except ValueError:
                            date_match = any(search_date in str(cell) for cell in row)
                    
                    if search_job:
                        job_match = any(search_job.lower() in str(cell).lower() for cell in row)
                    
                    return date_match and job_match
                
                filtered_data = [row for row in data_rows if matches_row(row)]
                search_term = f"Date: {search_date or 'Any'}, Job: {search_job or 'Any'}"
            else:
                filtered_data = data_rows
                search_term = "All Data"
            return render(request, 'viewer/display.html', {'data': filtered_data, 'search_date': search_term, 'total_rows': len(filtered_data), 'debug_rows': data_rows[:3]})
        else:
            # Homepage
            return render(request, 'viewer/homepage.html', {'title': 'RQC Data'})
    except Exception as e:
        # If file not found or error, show error
        return render(request, 'viewer/error.html', {'error': str(e)})
