import os
import sys
import time
import shutil
import re
import warnings
import logging
import pandas as pd
import config
import pdf_service

from logging.handlers import RotatingFileHandler
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --- DYNAMIC EMAIL ROUTING ---
if config.ACTIVE_EMAIL_PROVIDER == 'GMAIL':
    from gmail_mailing_script import send_report_email 
elif config.ACTIVE_EMAIL_PROVIDER == 'BREVO':
    from brevo_mailing_script import send_report_email
else:
    raise ValueError("[ERROR] Invalid ACTIVE_EMAIL_PROVIDER in config.py. Must be 'GMAIL' or 'BREVO'.")


warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')



# ==========================================
# 1. ARCHITECTURE & LOGGING
# ==========================================
log_formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
log_handler = RotatingFileHandler(
    config.LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
)
log_handler.setFormatter(log_formatter)

logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(log_handler)

# Redirect standard print() and crash errors to the logger silently
class StreamToLogger:
    def __init__(self, logger, log_level=logging.INFO):
        self.logger = logger
        self.log_level = log_level

    def write(self, buf):
        for line in buf.rstrip().splitlines():
            self.logger.log(self.log_level, line.rstrip())

    def flush(self):
        pass


sys.stdout = StreamToLogger(logger, logging.INFO)
sys.stderr = StreamToLogger(logger, logging.ERROR)


# ==========================================
# 1. EXCEL GENERATION LOGIC
# ==========================================
def generate_excel_report(m_p, m_l, m_inv, m_cur, m_ret, m_pct, g_total, dep_cc, h_deposit, 
                          o_p, o_l, o_inv, o_cur, o_ret, o_pct, 
                          c_p, c_l, c_inv, c_cur, c_ret, c_pct, alloc_data, report_date, report_folder):
    
    print(f"Generating enhanced multi-sheet Excel report for {report_date}...")
    
    date_str_iso = report_date.strftime("%Y-%m-%d")
    date_str_eu = report_date.strftime("%d.%m.%Y")
    
    # Save directly into the date-stamped folder
    output_filename = os.path.join(report_folder, f"Balancesheet_{date_str_iso}.xlsx")
    
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- Format Definitions ---
        font = 'Aptos Display'
        fmt_title = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'font_size': 14, 'font_name': font})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#008060', 'font_color': 'white', 'align': 'center', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        
        # Standard Text & Currency 
        fmt_even_txt = workbook.add_format({'bg_color': '#F2F2F2', 'align': 'left', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_odd_txt  = workbook.add_format({'bg_color': '#FFFFFF', 'align': 'left', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_even_cur = workbook.add_format({'bg_color': '#F2F2F2', 'align': 'center', 'num_format': '₹ #,##0.00', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_odd_cur  = workbook.add_format({'bg_color': '#FFFFFF', 'align': 'center', 'num_format': '₹ #,##0.00', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_even_pct = workbook.add_format({'bg_color': '#F2F2F2', 'align': 'center', 'num_format': '0.00%', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_odd_pct  = workbook.add_format({'bg_color': '#FFFFFF', 'align': 'center', 'num_format': '0.00%', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        
        # Color-Coded Currency & Percentages (Green/Red)
        fmt_even_cur_clr = workbook.add_format({'bg_color': '#F2F2F2', 'align': 'center', 'num_format': '[Color 10]₹ #,##0.00;[Red]₹ -#,##0.00;₹ 0.00', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_odd_cur_clr  = workbook.add_format({'bg_color': '#FFFFFF', 'align': 'center', 'num_format': '[Color 10]₹ #,##0.00;[Red]₹ -#,##0.00;₹ 0.00', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_even_pct_clr = workbook.add_format({'bg_color': '#F2F2F2', 'align': 'center', 'num_format': '[Color 10]0.00%;[Red]-0.00%;0.00%', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_odd_pct_clr  = workbook.add_format({'bg_color': '#FFFFFF', 'align': 'center', 'num_format': '[Color 10]0.00%;[Red]-0.00%;0.00%', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})

        # Bold Totals
        fmt_bold_txt = workbook.add_format({'bold': True, 'bg_color': '#EAEAEA', 'align': 'left', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_bold_cur = workbook.add_format({'bold': True, 'bg_color': '#EAEAEA', 'align': 'center', 'num_format': '₹ #,##0.00', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_bold_pct = workbook.add_format({'bold': True, 'bg_color': '#EAEAEA', 'align': 'center', 'num_format': '0.00%', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_bold_cur_clr = workbook.add_format({'bold': True, 'bg_color': '#EAEAEA', 'align': 'center', 'num_format': '[Color 10]₹ #,##0.00;[Red]₹ -#,##0.00;₹ 0.00', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        fmt_bold_pct_clr = workbook.add_format({'bold': True, 'bg_color': '#EAEAEA', 'align': 'center', 'num_format': '[Color 10]0.00%;[Red]-0.00%;0.00%', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})
        
        fmt_balance_title = workbook.add_format({'bold': True, 'bg_color': '#008060', 'font_color': 'white', 'align': 'center', 'border': 1, 'border_color': '#D3D3D3', 'font_name': font})

        # --- Reusable Engine to Draw a Standardized Sheet ---
        def build_sheet(worksheet, profit_df, loss_df, t_invested, t_current, t_returns, t_pct, include_balances=False, allocation=None):
            worksheet.hide_gridlines(2)
            
            # Auto-adjust column width based on the longest fund name
            all_names = pd.concat([profit_df, loss_df])['Scheme'].tolist() if not profit_df.empty or not loss_df.empty else []
            max_name_len = max([len(str(n)) for n in all_names]) if all_names else 35
            
            worksheet.set_column('A:A', 3)
            worksheet.set_column('B:B', max(35, max_name_len + 3)) # Dynamic width
            worksheet.set_column('C:G', 20) # Made wider for the extra Weight column
            
            # --- NEW: PDF / PRINT LAYOUT LOCK ---
            worksheet.set_landscape()
            worksheet.fit_to_pages(1, 0) # 1 page wide, infinite pages tall
            worksheet.set_margins(left=0.4, right=0.4, top=0.5, bottom=0.5)

            row_idx = 0
            
            # 1. HOLDING SUMMARY
            worksheet.merge_range(row_idx, 1, row_idx, 4, 'HOLDING SUMMARY', fmt_title)
            row_idx += 1
            
            headers_summary = ['Total Investments', 'Current Portfolio Value', 'Profit/Loss', 'Profit/Loss %']
            for col_num, header in enumerate(headers_summary):
                worksheet.write(row_idx, col_num + 1, header, fmt_header)
            row_idx += 1
            
            worksheet.write(row_idx, 1, t_invested, fmt_even_cur)
            worksheet.write(row_idx, 2, t_current, fmt_even_cur)
            worksheet.write(row_idx, 3, t_returns, fmt_even_cur_clr) # Colored
            worksheet.write(row_idx, 4, t_pct / 100, fmt_even_pct_clr) # Colored
            row_idx += 3 
            
            # 1.5 CATEGORY ALLOCATION (Only drawn if allocation data is passed)
            if allocation:
                worksheet.merge_range(row_idx, 1, row_idx, 3, 'CATEGORY ALLOCATION', fmt_title)
                row_idx += 1
                for col_num, header in enumerate(['Category', 'Total Value', 'Weight']):
                    worksheet.write(row_idx, col_num + 1, header, fmt_header)
                row_idx += 1
                
                for i, (cat_name, cat_val, cat_weight) in enumerate(allocation):
                    is_even = (i % 2 == 0)
                    worksheet.write(row_idx, 1, cat_name, fmt_even_txt if is_even else fmt_odd_txt)
                    worksheet.write(row_idx, 2, cat_val, fmt_even_cur if is_even else fmt_odd_cur)
                    worksheet.write(row_idx, 3, cat_weight, fmt_even_pct if is_even else fmt_odd_pct)
                    row_idx += 1
                row_idx += 2
            
            # HELPER: DRAW MUTUAL FUND TABLES
            def write_mf_table(start_row, title, df):
                worksheet.merge_range(start_row, 1, start_row, 6, title, fmt_title)
                start_row += 1
                
                # Added 'Weight' header
                headers = ['Scheme Name', 'Invested Value', 'Current Value', 'Returns', 'Percentage', 'Weight']
                for col_num, header in enumerate(headers):
                    worksheet.write(start_row, col_num + 1, header, fmt_header)
                start_row += 1
                
                for i, (_, row) in enumerate(df.iterrows()):
                    is_even = (i % 2 == 0)
                    
                    # Calculate weight relative to this specific sheet's total
                    weight = (row['Current_Value'] / t_current) if t_current > 0 else 0
                    
                    worksheet.write(start_row, 1, row['Scheme'], fmt_even_txt if is_even else fmt_odd_txt)
                    worksheet.write(start_row, 2, row['Invested_Value'], fmt_even_cur if is_even else fmt_odd_cur)
                    worksheet.write(start_row, 3, row['Current_Value'], fmt_even_cur if is_even else fmt_odd_cur)
                    worksheet.write(start_row, 4, row['Returns'], fmt_even_cur_clr if is_even else fmt_odd_cur_clr)
                    worksheet.write(start_row, 5, row['Percentage'] / 100, fmt_even_pct_clr if is_even else fmt_odd_pct_clr)
                    worksheet.write(start_row, 6, weight, fmt_even_pct if is_even else fmt_odd_pct)
                    start_row += 1
                    
                is_even = (len(df) % 2 == 0)
                for col in range(6):
                    if col == 0: fmt = fmt_even_txt if is_even else fmt_odd_txt
                    elif col == 4: fmt = fmt_even_pct_clr if is_even else fmt_odd_pct_clr
                    elif col == 5: fmt = fmt_even_pct if is_even else fmt_odd_pct
                    elif col == 3: fmt = fmt_even_cur_clr if is_even else fmt_odd_cur_clr
                    else: fmt = fmt_even_cur if is_even else fmt_odd_cur
                    worksheet.write(start_row, col + 1, "", fmt)
                start_row += 1
                
                tot_inv = df['Invested_Value'].sum()
                tot_cur = df['Current_Value'].sum()
                tot_ret = df['Returns'].sum()
                tot_pct = (tot_ret / tot_inv) if tot_inv != 0 else 0
                tot_weight = (tot_cur / t_current) if t_current > 0 else 0
                
                worksheet.write(start_row, 1, "Total", fmt_bold_txt)
                worksheet.write(start_row, 2, tot_inv, fmt_bold_cur)
                worksheet.write(start_row, 3, tot_cur, fmt_bold_cur)
                worksheet.write(start_row, 4, tot_ret, fmt_bold_cur_clr)
                worksheet.write(start_row, 5, tot_pct, fmt_bold_pct_clr)
                worksheet.write(start_row, 6, tot_weight, fmt_bold_pct)
                
                return start_row + 3 

            # 2. PROFIT & LOSS HOLDINGS
            if not profit_df.empty:
                row_idx = write_mf_table(row_idx, f"PROFIT HOLDINGS AS ON {date_str_iso}", profit_df)
            if not loss_df.empty:
                row_idx = write_mf_table(row_idx, f"LOSS HOLDINGS AS ON {date_str_iso}", loss_df)
                
            # 3. BALANCE SUMMARY (Only drawn on Liquid Funds)
            if include_balances:
                worksheet.merge_range(row_idx, 1, row_idx, 2, f"Balance as on {date_str_eu}", fmt_balance_title)
                row_idx += 1
                balances = [
                    ("Groww MF", t_current),
                    ("Deposit CC", dep_cc),
                    ("House Owner Deposit", h_deposit)
                ]
                for i, (label, val) in enumerate(balances):
                    is_even = (i % 2 != 0) 
                    worksheet.write(row_idx, 1, f"{label}:", fmt_even_txt if is_even else fmt_odd_txt)
                    worksheet.write(row_idx, 2, val, fmt_even_cur if is_even else fmt_odd_cur)
                    row_idx += 1
                    
                is_even = (len(balances) % 2 != 0)
                worksheet.write(row_idx, 1, "", fmt_even_txt if is_even else fmt_odd_txt)
                worksheet.write(row_idx, 2, "", fmt_even_cur if is_even else fmt_odd_cur)
                row_idx += 1
                worksheet.write(row_idx, 1, "Grand total:", fmt_bold_txt)
                worksheet.write(row_idx, 2, g_total, fmt_bold_cur)

        # 1. Create Sheet 1: Liquid Funds (Dad's Tracked)
        ws_main = workbook.add_worksheet('Liquid Funds')
        build_sheet(ws_main, m_p, m_l, m_inv, m_cur, m_ret, m_pct, include_balances=True)
        
        # 2. Create Sheet 2: Equity Funds
        if not o_p.empty or not o_l.empty:
            ws_other = workbook.add_worksheet('Equity Funds')
            build_sheet(ws_other, o_p, o_l, o_inv, o_cur, o_ret, o_pct, include_balances=False)

        # 3. Create Sheet 3: Consolidated Folio
        ws_consolidated = workbook.add_worksheet('Consolidated Folio')
        build_sheet(ws_consolidated, c_p, c_l, c_inv, c_cur, c_ret, c_pct, include_balances=False, allocation=alloc_data)

    print(f"[SUCCESS] Final multi-sheet formatted report saved to: {output_filename}")

# ==========================================
# 2. TRANSFORMATION LOGIC (pandas)
# ==========================================
def process_report(filepath):
    print(f"Processing newly detected report: {filepath}")
    
    try:
        raw_df = pd.read_excel(filepath, header=None)
        
        report_date_str = None
        for index, row in raw_df.iterrows():
            for val in row.dropna():
                if isinstance(val, str) and 'HOLDINGS AS ON' in val.upper():
                    match = re.search(r'HOLDINGS AS ON\s+([0-9\-]+)', val, re.IGNORECASE)
                    if match:
                        report_date_str = match.group(1)
                        break
            if report_date_str:
                break
                
        if not report_date_str:
            print("[ERROR] Could not find 'HOLDINGS AS ON' date inside the Excel file. Aborting transformation.")
            return 
            
        report_date = pd.to_datetime(report_date_str).date()
        print(f"Report actual date identified as: {report_date}")

        header_row_index = raw_df[raw_df.eq('Scheme Name').any(axis=1)].index[0]
        df = pd.read_excel(filepath, header=header_row_index)
        
        df.columns = df.columns.str.strip().str.replace('\n', '')
        df = df.dropna(subset=['Scheme Name'])
        
        df = df.rename(columns={
            'Scheme Name': 'Scheme',
            'Invested Value': 'Invested_Value',
            'Current Value': 'Current_Value'
        })
        
        if 'Current_Value' not in df.columns:
            print(f"[DEBUG] Column renaming failed! Exact headers found: {df.columns.tolist()}")
            return  
            
        df['Invested_Value'] = pd.to_numeric(df['Invested_Value'], errors='coerce')
        df['Current_Value'] = pd.to_numeric(df['Current_Value'], errors='coerce')
        
        df['Returns'] = df['Current_Value'] - df['Invested_Value']
        df['Percentage'] = (df['Returns'] / df['Invested_Value']) * 100

        
        # Partition data streams
        main_df = df[~df['Scheme'].isin(config.SEPARATED_FUNDS)]
        other_df = df[df['Scheme'].isin(config.SEPARATED_FUNDS)]
        
        def get_metrics(data_df):
            p_df = data_df[data_df['Returns'] >= 0].sort_values(by='Percentage', ascending=False)
            l_df = data_df[data_df['Returns'] < 0].sort_values(by='Percentage', ascending=True)
            t_inv = data_df['Invested_Value'].sum()
            t_cur = data_df['Current_Value'].sum()
            t_ret = t_cur - t_inv
            t_pct = (t_ret / t_inv * 100) if t_inv else 0
            return p_df, l_df, t_inv, t_cur, t_ret, t_pct
            
        m_p, m_l, m_inv, m_cur, m_ret, m_pct = get_metrics(main_df)     # Liquid Metrics
        o_p, o_l, o_inv, o_cur, o_ret, o_pct = get_metrics(other_df)    # Equity Metrics
        c_p, c_l, c_inv, c_cur, c_ret, c_pct = get_metrics(df)          # Consolidated Metrics
        
        # Calculate Category Allocation for the Consolidated Sheet
        alloc_data = [
            ("Liquid Funds", m_cur, (m_cur / c_cur) if c_cur > 0 else 0),
            ("Equity Funds", o_cur, (o_cur / c_cur) if c_cur > 0 else 0)
        ]
        
        grand_total = m_cur + config.DEPOSIT_CC + config.HOUSE_DEPOSIT  
        
        # --- NEW: FOLDER STRUCTURE LOGIC ---
        date_str_iso = report_date.strftime('%Y-%m-%d')
        report_folder = os.path.join(config.REPORTS_BASE_FOLDER, date_str_iso)
        os.makedirs(report_folder, exist_ok=True)
        
        # Generate the Excel File inside the new folder
        generate_excel_report(
            m_p, m_l, m_inv, m_cur, m_ret, m_pct, grand_total, config.DEPOSIT_CC, config.HOUSE_DEPOSIT,
            o_p, o_l, o_inv, o_cur, o_ret, o_pct,
            c_p, c_l, c_inv, c_cur, c_ret, c_pct, alloc_data,
            report_date, report_folder
        )
        
        excel_filename = os.path.join(report_folder, f"Balancesheet_{date_str_iso}.xlsx")
        pdf_filename = os.path.join(report_folder, f"Balancesheet_{date_str_iso}.pdf")

        # Generate PDF
        pdf_service.generate_pdf(excel_filename, pdf_filename)

        # Move the downloaded file itself (with its original name) into the date folder
        original_filename = os.path.basename(filepath)
        final_raw_path = os.path.join(report_folder, original_filename)
        
        try:
            # shutil.move physically moves the file, ensuring no copy is left in Downloads
            shutil.move(filepath, final_raw_path)
            print(f"[CLEANUP] Successfully moved original download to: {final_raw_path}")
        except Exception as e:
            print(f"[ERROR] Could not move original file to {final_raw_path}: {e}")

        # --- EMAIL TRIGGER ENABLED ---
        send_report_email(excel_filename, pdf_filename, report_date)
        
    except Exception as e:
        print(f"Transformation Failed: {e}")

# ==========================================
# 3. FOLDER MONITORING LOGIC (watchdog)
# ==========================================
def wait_for_file_unlock(filepath, timeout=30):
    start_time = time.time()
    while True:
        try:
            with open(filepath, 'a'):
                pass
            if os.path.getsize(filepath) > 0:
                return True
        except (PermissionError, IOError):
            pass 
            
        if time.time() - start_time > timeout:
            print(f"[ERROR] Timeout: Browser never released the file: {filepath}")
            return False
        time.sleep(1)

class GrowwReportHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory:
            self.check_and_process(event.src_path)
            
    def on_moved(self, event):
        if not event.is_directory:
            self.check_and_process(event.dest_path)
            
    def check_and_process(self, filepath):
        filename = os.path.basename(filepath)
        clean_name = filename.lower().replace(" ", "_").replace("-", "_")
        
        if filepath.lower().endswith('.xlsx') and clean_name.startswith("mutual_fund"):
            
            if not os.path.exists(filepath):
                return
                
            print(f"\n[TRIGGER] Groww report detected: {filename}")
            
            if not wait_for_file_unlock(filepath):
                return 
            
            process_report(filepath)

if __name__ == "__main__":
    event_handler = GrowwReportHandler()
    observer = Observer()
    observer.schedule(event_handler, config.WATCH_FOLDER, recursive=False)
    observer.start()
    print(f"Service Started: Monitoring {config.WATCH_FOLDER} for new files...")
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()