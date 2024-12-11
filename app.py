from flask import Flask,render_template,request,flash,redirect,url_for
import win32com.client as win32
import os
from datetime import datetime
import pythoncom
import traceback
import win32com
import shutil
import pywintypes


app = Flask(__name__)

app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

@app.route('/')
def index():

    return render_template("index.html")


@app.route('/start_print', methods=['POST'])
def start_print():
    try:

        excel = None
        workbook = None
          
        file = request.files['file']
        start_sheet_number = request.form.get("start_sheet_number",None)
        end_sheet_number = request.form.get("end_sheet_number",None)


        if file.filename == '':
            flash("No file selected! Please Choose File","danger")
            return redirect(url_for('index'))

        if not file.filename.split(".")[-1] == 'xlsx':
            flash("Please Choose Excle File","danger")
            return redirect(url_for('index'))
        
        if not start_sheet_number  or not end_sheet_number:
            flash("Please Enter Start And End Sheet Number! It Cannot Be Empty","danger")
            return redirect(url_for('index'))
        
        # if not isinstance( start_sheet_number,int)  or not isinstance(end_sheet_number,int):
        #     flash("Please Enter Start And End Sheet Number! It Cannot Be Alphabhates","danger")
        #     return redirect(url_for('index'))
        
        
        file.save(file.filename)

        # Initialize the COM library
        pythoncom.CoInitialize()

        # Create a new Excel application instance
        excel = win32.gencache.EnsureDispatch('Excel.Application')

        excel_file = file.filename

        # Open the Excel workbook
        workbook = excel.Workbooks.Open(os.path.abspath(excel_file))

        start_sheet_number = int(start_sheet_number)
        end_sheet_number = int(end_sheet_number)

        for sheet in range(start_sheet_number,end_sheet_number + 1):

            sheet_name = f"GT{sheet}"

            # Access the desired worksheet
            sheet = workbook.Sheets(sheet_name)

            # Access and read specific cells (e.g., D6 and D7)
            custome_name = sheet.Range('D6').Value
            date = sheet.Range('G7').Value

            # Check the type of 'date' and convert it if necessary
            if isinstance(date, datetime):
                dt_obj = date
                # Convert the datetime object to the desired format
                date_formatted = dt_obj.strftime('%d-%m-%Y')
            elif isinstance(date, str):
                try:
                    # Attempt to parse the date string in the 'd-m-y' format
                    date_formatted = datetime.strptime(date, '%d-%m-%Y')
                except ValueError:
                    date_formatted = date.replace('/','-')



            print_area = 'B3:H36'

            sheet.PageSetup.PaperSize = win32.constants.xlPaperA4
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.FitToPagesTall = 1
            sheet.PageSetup.FitToPagesWide = 1
            sheet.PageSetup.PrintArea = print_area

            if not os.path.exists('Chetna_Plastic_Bills'):
                os.mkdir('Chetna_Plastic_Bills')

            # Print the worksheet to PDF
            pdf_file = f'C:\\Users\\vader\\Downloads\\flask_excel_print-main\\flask_excel_print-main\\Chetna_Plastic_Bills\\{custome_name} ({date_formatted}).pdf'
            sheet.ExportAsFixedFormat(0, pdf_file)

            print(f"PDF saved as '{pdf_file}'")

          

        flash("Your PDFs Downloaded Successfully","success")
        return render_template("index.html")
    
    except Exception as e:
        traceback.print_exc()  # Print the traceback information
        return str(e)
    
    finally:
        try: 
            if workbook:
                workbook.Close(False)
        except Exception as close_exception:
            print(f"Error while closing the workbook: {str(close_exception)}")

        try:
            if excel:
                excel.Quit()
        except Exception as quit_exception:
            print(f"Error while quitting Excel: {str(quit_exception)}")
        try:
            
            os.remove(file.filename)
            
        except Exception as quit_exception:
            print(f"Error while quitting Excel: {str(quit_exception)}")

        try:
                # Get the directory specified by win32com.__gen_path__
                gen_path = win32com.__gen_path__

                # Check if the directory exists
                if os.path.exists(gen_path):
                    # List all files in the directory
                    for file_name in os.listdir(gen_path):
                        file_path = os.path.join(gen_path, file_name)
                        if os.path.isfile(file_path):
                            # Delete each file in the directory
                            os.remove(file_path)
                            print(f"Deleted file: {file_path}")
        except Exception as delete_exception:
            print(f"Error while deleting files in the directory: {str(delete_exception)}")



def clear_bill_folder():
    folder_path = f"Chetna_Plastic_Bills"

    if os.path.exists(folder_path):

        for file_name in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file_name)
            if os.path.isfile(file_path):
                # Delete each file in the directory
                os.remove(file_path)
                print(f"Deleted PDF file: {file_path}")








if __name__ == '__main__':
    clear_bill_folder()
    app.run(debug=True)

# https://stackoverflow.com/questions/52889704/python-win32com-excel-com-model-started-generating-errors

# Deleting the gen_py output directory and re-running makepy SUCCEEDS and subsequently the test application runs OK again.

# So the symptom is resolved, but any clues as to how this could have happened. This is a VERY long running application (think 24x7 for years) and I'm concerned that whatever caused this might occur again.

# To find the output directory, run this in your python console / python session:

# import win32com
# print(win32com.__gen_path__)
# or, even better, a one-liner in the command line:

# python -c "import win32com; print(win32com.__gen_path__)"

# Based on the exception message in your post, the directory you need to remove will be titled '00020813-0000-0000-C000-000000000046x0x1x9'. So delete this directory and re-run the code. And if you're nervous about deleting it (like I was) just cut the directory and paste it somewhere else.

# 💡Note that this directory is usually in your "TEMP" directory (copy-paste %TEMP%/gen_py in Windows File Explorer and you will arrive there directly).

# I have no idea why this happens nor do I know how to prevent it from happening again, but the directions in the link I provided seemed to work for me.
