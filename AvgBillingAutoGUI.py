import PySimpleGUI as sg
import datetime
import pandas as pd
import xlsxwriter
import os

sg.theme('Dark Blue 3')

# All the stuff inside your window.
layout = [  
            [sg.Text('Add school invoice sheet', size=(20, 1)), sg.Input(size=(53, 1)), sg.FileBrowse('File Browse', size=(12, 1))],
            [sg.Text('Add school CSV file folder', size=(20, 1)), sg.Input(key='foldername', size=(53, 1)), sg.Button('Folder Browse', size=(12, 1))],
            [sg.Multiline(key='files', size=(92,5), autoscroll=True, auto_refresh=True)],
            [sg.Text('Elapsed Time:'), sg.Output(key='elapsed_time', size=(20,1))],
            [sg.Button('Start'), sg.Button('Cancel')],
        ]

# Create the Window
window = sg.Window('Automated Average Student Billing - Version 1.0', layout)
# Event Loop to procepyinstallerss "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    if event == 'Folder Browse':
        global foldername
        foldername = sg.PopupGetFolder('Select folder', no_window=True)
        if foldername: # `None` when clicked `Cancel` - so I skip it
            window['foldername'].update(foldername)
            filenames = sorted(os.listdir(foldername))
            # it use `key='files'` to `Multiline` widget
            window['files'].update("\n".join(filenames))
    if event == 'Start':
        class AvgStudentBillingAutomation:
            def __init__(self, customer, invoice, cost, invoice_date_adj, date_style, name_csv_file, covid_charge):
                customer = self.customer = customer
                invoice = self.invoice = invoice
                cost = self.cost = float(cost)
                invoice_date_adj = self.invoice_date_adj = invoice_date_adj
                date_style = self.date_style = date_style
                name_csv_file = self.name_csv_file = name_csv_file
                covid_charge = self.covid_charge = covid_charge

                #folder_name = values[1]

            
                # date of invoice
                current_date = datetime.datetime.now()
                if invoice_date_adj == 0:
                    if date_style == "mm_yyyy":
                        invoice_date = "{}_{}".format(current_date.year, current_date.month)
                    elif date_style == "mm_dd_yyyy":
                        invoice_date = "{}_{}_{}".format(current_date.month, current_date.day, current_date.year)
                elif invoice_date_adj == -1:
                    if current_date.month == 1:
                        invoice_date = "{}_{}".format(12, current_date.year - 1)
                    else:
                        invoice_date = "{}_{}".format(current_date.month - 1, current_date.year)

            
                """ Open and read the CSV file of data """
                col_names = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']

                # determine if folder 'Complete' exists, if not, create it
                if not os.path.exists('../2 - Complete - 2/{}_{}'.format(current_date.year, current_date.month)):
                    os.makedirs('../2 - Complete - 2/{}_{}'.format(current_date.year, current_date.month))
                try:
                    df = pd.read_csv(f"{foldername}/{name_csv_file}.csv", names=col_names, keep_default_na=False)


            
                    """ Create Excel sheet to save in customers folder """
                    # these two lines generate the workbook/worksheet to write to
                    workbook = xlsxwriter.Workbook(('../2 - Complete - 2/{}_{}/{} Invoice Attachment_{}.xlsx').format(current_date.year, current_date.month, customer, invoice_date))
                    workbook.formats[0].set_font_size(10)
                    worksheet = workbook.add_worksheet()

                    # cell formatting
                    page_name_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 14, 'align': 'center'})
                    metadata_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'center', 'underline': 1})
                    header_cell_formatL = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'left', 'bold': True})
                    header_cell_formatR = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'right', 'bold': True})
                    date_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'left'})
                    count_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 8, 'align': 'right'})
                    bottom_text_cell_format = workbook.add_format({'font_name': 'Verdana', 'font_size': 7.5, 'align': 'center', 'bold': True})


                    # df to human readable variables
                    if df.loc[1, 'E'] == 'LOCATION:':
                        report_dates = df.loc[1,'C']
                        location_id = df.loc[1,'F']
                        avg_num_students = float(df.loc[1,'H'])
                        dates_count_dict = pd.Series(df.J.values,index=df.I).to_dict()

                        report_total = avg_num_students * cost

                    elif df.loc[1, 'D'] == 'LOCATION:':
                        report_dates = df.loc[1,'B']
                        location_id = df.loc[1,'E']
                        avg_num_students = float(df.loc[1,'G'])
                        dates_count_dict = pd.Series(df.I.values,index=df.H).to_dict()
                        
                        report_total = avg_num_students * cost
                    
                    # report total calculation
                    #report_total = avg_num_students * cost

                    """ Set column widths """
                    worksheet.set_column('A:A', 20)
                    worksheet.set_column('B:B', 17)
                    worksheet.set_column('C:C', 30)
                    worksheet.set_column('D:D', 10)
                    worksheet.set_column('E:E', 20)

                    """ Heading, Metadata, and Footer for the sheet and Merge cells """
                    worksheet.merge_range('A1:E1', 'BILLING AVERAGE STUDENTS', page_name_cell_format)
                    worksheet.merge_range('A2:E2', 'CUSTOMER: {}'.format(customer), metadata_cell_format)
                    worksheet.merge_range('A3:E3', 'LOCATION: {}'.format(location_id), metadata_cell_format)
                    worksheet.merge_range('A4:E4', 'INVOICE: {}'.format(invoice), metadata_cell_format)
                    worksheet.merge_range('A5:E5', report_dates, metadata_cell_format)
                    worksheet.write('B6', 'LOCATION: {}'.format(location_id), header_cell_formatL)
                    worksheet.write('C6', 'AVERAGE STUDENTS:', header_cell_formatR)
                    worksheet.write('D6', avg_num_students, header_cell_formatR)

                    """ Add CSV file data to the Excel sheet """
                    row = 6
                    col = 1

                    for day, count in dates_count_dict.items():
                        worksheet.write(row, col, day, date_cell_format)
                        worksheet.write(row, col + 2, count, count_cell_format)
                        row += 1        
                    
                    """ Add report totals to sheets """
                    worksheet.merge_range(f'A{row + 1}:E{row + 1}', f"CALCULATED COST OF {location_id} LOCATION STUDENT AVERAGE OF {avg_num_students} AT ${'{0:.2f}'.format(cost)}/MONTH IS ${'{0:.2f}'.format(report_total)}", bottom_text_cell_format)
                        #.format(location_id, avg_num_students, '{0:.2f}'.format(cost), '{0:.2f}'.format(report_total)), bottom_text_cell_format)
                    row += 1

                    if report_total < 500:
                        worksheet.merge_range(f'A{row + 1}:E{row + 1}', 'MINIMUM AMOUNT DUE AS PER AGREEMENT IS $500.00', bottom_text_cell_format)
                    row += 1

                    try:
                        if covid_charge.upper() == 'YES':
                            worksheet.merge_range(f'A{row + 1}:E{row + 1}', 'DURING COVID-19 MINIMUM AMOUNT DUE AS PER AGREEMENT IS $500.00', bottom_text_cell_format)
                    except AttributeError:
                        pass


                    workbook.close()
                    #os.remove("./6 - ASB CSV Files - 6/{}.csv".format(name_csv_file))
                
                except FileNotFoundError:
                    with open('../2 - Complete - 2/{}_{}/Missing CSV Files - {}_{}.txt'.format(current_date.year, current_date.month, current_date.year, current_date.month), 'a') as file:
                        file.write('{} - Missing CSV file\n'.format(customer))


        def billing_instances():
            ''' This function takes an Excel spreadsheet with the schools and 
            their variables, reads it, then makes it understandable by the above 
            Class BillingAutomation to itterate through the schools'''
            
            """ Open and read Excel sheet of schools """
            schools_file = pd.read_excel(values[0], sheet_name='Sheet1', engine='openpyxl')


            schools = schools_file.values.tolist()

            billingautomation_instances = []
            for school in schools:
                billingautomation_instances.append(AvgStudentBillingAutomation(*school))


        start = datetime.datetime.now()
        print('Automated Average Student Billing Start')
        billing_instances()
        window['elapsed_time'].update(datetime.datetime.now() - start)
        

window.close()