import os
import logging
import win32com.client
import pythoncom

logger = logging.getLogger(__name__)

def generate_pdf(excel_path, pdf_path):
    logger.info(f"Generating PDF for: {os.path.basename(excel_path)}")
    
    if not os.path.exists(excel_path):
        logger.error("[ERROR] Excel file not found for PDF conversion.")
        return None
    excel = None
    wb = None
    
    try:
        # Crucial for Watchdog: Initializes COM for background threads
        pythoncom.CoInitialize() 
        
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # Open the generated Excel file
        wb = excel.Workbooks.Open(excel_path)
        
        # Select all active worksheets so they print as a single continuous document
        wb.Worksheets.Select()
        
        # Export parameter 0 = xlTypePDF
        wb.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
        
        logger.info(f"[SUCCESS] PDF generated: {os.path.basename(pdf_path)}")
        return pdf_path
        
    except Exception as e:
        logger.error(f"[ERROR] PDF Generation failed: {e}")
        return None
    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()