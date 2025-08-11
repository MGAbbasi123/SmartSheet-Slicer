✂️ Excel Sheet Splitter: Automate Your Data Organization
The Excel Sheet Splitter is a powerful and intuitive desktop application built with CustomTkinter, designed to transform large Excel workbooks into highly organized, multi-sheet files. Say goodbye to tedious manual data separation! This tool automates the process of splitting your Excel data based on unique values in a specified column, creating a new, dedicated sheet for each unique entry.

✨ Key Features
User-Friendly Interface: A clean, modern graphical interface makes file selection and column input straightforward.

Dynamic Splitting: Split your Excel file by any specified column, adapting to your specific data organization needs.

Automated Sheet Creation: Generates a distinct Excel sheet for each unique value found in your chosen column.

Intelligent Indexing: Creates a master "Index" sheet with active hyperlinks that allow you to jump directly to any split sheet with a single click. Each split sheet also includes a "Back to Index" link.

Robust Sheet Naming: Automatically cleans invalid characters and handles duplicate sheet names by ensuring uniqueness.

Responsive Performance: Built with multithreading, the application remains responsive during processing, even for very large Excel files.

Error Handling: Provides clear messages for issues like missing files or incorrect column names.

Customizable Output: Allows you to specify the name of the resulting combined Excel workbook.

🛠️ Prerequisites
To run this application, you need Python 3.x installed on your computer, along with the customtkinter, pandas, and openpyxl libraries.

You can install the required libraries using pip from your terminal or command prompt:

pip install customtkinter pandas openpyxl

🚀 How to Run the Application
Save the Script:
Copy the entire Python code for the Excel Sheet Splitter application (from the excel-splitter-ui-app-fixed section in our previous conversation) and save it to a file named excel_splitter_app.py (or any .py name you prefer) on your local computer.

Open Your Terminal/Command Prompt:
Navigate to the directory where you saved excel_splitter_app.py.

cd /path/to/your/script/directory

Execute the Script:
Run the script using your Python interpreter:

python excel_splitter_app.py

A new desktop window titled "Excel Splitter with Index" should appear.

⚙️ How to Use the App
Select Source Excel File:

Click the "Browse" button next to "Source Excel File:".

A file dialog will open. Navigate to and select the .xlsx or .xls file you want to split. The selected path will appear in the input field.

Enter Column Name to Split By:

In the "Column to Split By:" field, type the exact name of the column in your Excel file that contains the values you want to use for splitting (e.g., ER_NO, CustomerID, Department). Case and spacing must match exactly the column header in your Excel file.

Enter Output File Name (Optional):

You can provide a desired name for the output Excel file (e.g., "My_Split_Excel"). If left blank, it will default to a name based on your source file (e.g., YourFile_Split.xlsx).

Split & Create Workbook:

Click the "⚙️ Split & Create Workbook" button.

The app will process the file. A "Processing..." message will appear on the button, and the status bar will show the progress.

Upon completion, a success message and the path to the newly created Excel workbook will be displayed, along with a pop-up confirmation.

💡 Troubleshooting
ModuleNotFoundError: No module named 'customtkinter' (or pandas, openpyxl)
This error means a required library is not installed in your Python environment.

Solution: Open your terminal or command prompt and run:
pip install customtkinter pandas openpyxl
After installation, try running the application again.

_tkinter.TclError: no display name and no $DISPLAY environment variable
This error occurs when you try to run a graphical application (like this CustomTkinter app) in an environment without a graphical display server (e.g., a remote server, a headless VM, or certain online coding platforms).

Solution: This application is designed to run directly on your local computer's operating system (Windows, macOS, or Linux with a desktop environment). Ensure you are running the excel_splitter_app.py script from your local machine's terminal or command prompt.

Column Not Found Error (Pop-up Message)
This error indicates that the "Column to Split By" name you entered does not exactly match any column header in your Excel file.

Solution:

Open your source Excel file.

Carefully check the spelling, casing (e.g., ER_NO vs er_no), and any spaces (e.g., ER No vs ER_NO) of the column header you intend to use.

Enter the exact matching name into the application's "Column to Split By:" field.

Created by MG