import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os
import math

def excel_sheets_to_multi_page_pdf(excel_path, output_folder, rows_per_page=30):
    # Load Excel file
    xl = pd.ExcelFile(excel_path)
    
    # Create output folder , if you want you can give your oun output folder address
    os.makedirs(output_folder, exist_ok=True)

    # Loop through each sheet
    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name)
        total_rows = len(df)
        total_pages = math.ceil(total_rows / rows_per_page)

        # Create a new PDF for each sheet
        pdf_filename = os.path.join(output_folder, f"{sheet_name}.pdf")
        with PdfPages(pdf_filename) as pdf:
            for page in range(total_pages):
                start = page * rows_per_page
                end = min(start + rows_per_page, total_rows)
                df_page = df.iloc[start:end]

                fig, ax = plt.subplots(figsize=(20, rows_per_page * 0.3 + 1))
                ax.axis('tight')
                ax.axis('off')
                table = ax.table(cellText=df_page.values,
                                 colLabels=df.columns,
                                 cellLoc='center',
                                 loc='center')
                # Remove borders from all cells
                for key, cell in table.get_celld().items():
                    cell.set_linewidth(0)


                ax.set_title(f'{sheet_name} - Page {page + 1} of {total_pages}', fontsize=12, pad=10)
                pdf.savefig(fig, bbox_inches='tight')
                plt.close()
    
    print(f"All sheets saved as paginated PDFs in '{output_folder}'")

# here you can  provide the info
excel_file = r"C:\Users\DELL\OneDrive\Desktop\Internshipprojects\converting_excelsheet_topdf\excel_dataset\ExcelTestData1.xlsx"  # Replace with your Excel file path
output_dir = "output_pdfs"
excel_sheets_to_multi_page_pdf(excel_file, output_dir, rows_per_page=30)
