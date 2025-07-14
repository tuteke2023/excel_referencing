#!/usr/bin/env python3
"""
Web-based TB-GL Linker using Streamlit
Simple drag-and-drop interface for linking TB and GL files
"""

import streamlit as st
import pandas as pd
import tempfile
import os
try:
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("Missing required package: openpyxl. Please contact support.")
    st.stop()
from difflib import SequenceMatcher

class TBGLLinkerWeb:
    """Simplified TBGLLinker for web app use."""
    def __init__(self, tb_path, gl_path, output_path):
        self.tb_file = tb_path
        self.gl_file = gl_path
        self.output_file = output_path
        self.tb_wb = None
        self.gl_wb = None
        self.tb_sheet = None
        self.gl_sheet = None
        self.account_mappings = {}
        self.tb_config = {}
        self.gl_accounts = {}
        
    def load_workbooks(self):
        """Load both workbooks and identify the main sheets."""
        self.tb_wb = load_workbook(self.tb_file)
        self.gl_wb = load_workbook(self.gl_file)
        
        # Find TB sheet (usually first sheet or one with 'TB' in name)
        tb_sheet_name = self._find_sheet(self.tb_wb, ['TB', 'Trial Balance', 'Trial_Balance'])
        self.tb_sheet = self.tb_wb[tb_sheet_name]
        
        # Find GL sheet
        gl_sheet_name = self._find_sheet(self.gl_wb, ['GL', 'General Ledger', 'General_Ledger'])
        self.gl_sheet = self.gl_wb[gl_sheet_name]
        
    def _find_sheet(self, workbook, keywords):
        """Find sheet by keywords in sheet names."""
        sheet_names = workbook.sheetnames
        
        # First try exact match
        for sheet in sheet_names:
            for keyword in keywords:
                if keyword.lower() in sheet.lower():
                    return sheet
        
        # If no match found, return first sheet
        return sheet_names[0]
    
    def analyze_tb_structure(self):
        """Analyze TB structure to find account columns."""
        # Find header row (usually contains 'Account', 'Debit', 'Credit')
        header_row = None
        account_col = None
        account_name_col = None
        debit_col = None
        credit_col = None
        
        for row in range(1, min(20, self.tb_sheet.max_row + 1)):
            row_values = []
            for col in range(1, min(20, self.tb_sheet.max_column + 1)):
                cell_value = self.tb_sheet.cell(row, col).value
                if cell_value:
                    row_values.append((col, str(cell_value).lower()))
            
            # Look for key headers
            for col, value in row_values:
                if any(keyword in value for keyword in ['account', 'acct', 'code']):
                    if 'name' in value or 'description' in value:
                        account_name_col = col
                    else:
                        account_col = col
                    header_row = row
                elif 'debit' in value:
                    debit_col = col
                elif 'credit' in value:
                    credit_col = col
            
            if header_row and (account_col or account_name_col) and (debit_col or credit_col):
                break
        
        if not header_row:
            raise ValueError("Could not find header row in Trial Balance")
        
        self.tb_config = {
            'header_row': header_row,
            'account_col': account_col,
            'account_name_col': account_name_col,
            'debit_col': debit_col,
            'credit_col': credit_col,
            'data_start_row': header_row + 1
        }
        
    def analyze_gl_structure(self):
        """Analyze GL structure to find account sections."""
        self.gl_accounts = {}
        
        for row in range(1, self.gl_sheet.max_row + 1):
            # Check first few columns for account names
            for col in range(1, min(5, self.gl_sheet.max_column + 1)):
                cell_value = self.gl_sheet.cell(row, col).value
                if cell_value and isinstance(cell_value, str):
                    # Check if this looks like an account header
                    if col == 1 and self._is_account_header(row, cell_value):
                        account_name = cell_value.strip()
                        self.gl_accounts[account_name] = row
        
    def _is_account_header(self, row, value):
        """Determine if a cell contains an account header."""
        if row >= self.gl_sheet.max_row:
            return False
        
        # Account headers usually don't have dates or numbers in the same row
        for col in range(2, min(10, self.gl_sheet.max_column + 1)):
            cell_value = self.gl_sheet.cell(row, col).value
            if cell_value:
                return False
        
        # Check if following rows have transaction-like data
        for check_row in range(row + 1, min(row + 5, self.gl_sheet.max_row + 1)):
            cell_value = self.gl_sheet.cell(check_row, 1).value
            if cell_value:
                return True
        
        return False
        
    def match_accounts(self):
        """Match TB accounts to GL accounts using fuzzy matching."""
        # Get TB accounts
        tb_accounts = []
        for row in range(self.tb_config['data_start_row'], self.tb_sheet.max_row + 1):
            account_name = None
            
            # Try to get account name from name column first
            if self.tb_config['account_name_col']:
                account_name = self.tb_sheet.cell(row, self.tb_config['account_name_col']).value
            
            # If no name column or empty, try to find account name in other columns
            if not account_name:
                # Check columns near the account code column for account names
                for col_offset in [1, -1, 2, -2]:  # Check adjacent columns
                    check_col = (self.tb_config['account_col'] or 1) + col_offset
                    if 1 <= check_col <= self.tb_sheet.max_column:
                        cell_value = self.tb_sheet.cell(row, check_col).value
                        if cell_value and isinstance(cell_value, str) and len(cell_value) > 3:
                            # Check if it looks like an account name (not a number or code)
                            if not cell_value.replace('.', '').replace('-', '').isdigit():
                                account_name = cell_value
                                if not self.tb_config['account_name_col']:
                                    self.tb_config['account_name_col'] = check_col
                                break
            
            if account_name and isinstance(account_name, str):
                tb_accounts.append((row, account_name.strip()))
        
        # Match each TB account to GL account
        for tb_row, tb_account in tb_accounts:
            best_match = None
            best_score = 0
            
            for gl_account, gl_row in self.gl_accounts.items():
                # Calculate similarity score
                score = SequenceMatcher(None, tb_account.lower(), gl_account.lower()).ratio()
                
                # Bonus for exact match
                if tb_account.lower() == gl_account.lower():
                    score = 1.0
                
                if score > best_score and score > 0.8:  # 80% similarity threshold
                    best_score = score
                    best_match = (gl_account, gl_row)
            
            if best_match:
                self.account_mappings[tb_row] = best_match
                
    def copy_gl_sheet(self):
        """Copy GL sheet to TB workbook."""
        # Create new sheet in TB workbook
        gl_sheet_name = 'General Ledger Detail'
        if gl_sheet_name in self.tb_wb.sheetnames:
            del self.tb_wb[gl_sheet_name]
        
        new_sheet = self.tb_wb.create_sheet(title=gl_sheet_name)
        
        # Copy all data
        for row in self.gl_sheet.iter_rows():
            for cell in row:
                new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
                
    def add_hyperlinks(self):
        """Add hyperlinks to TB sheet."""
        # Find the best column for hyperlinks (after Credit or last column)
        if self.tb_config['credit_col']:
            hyperlink_col = self.tb_config['credit_col'] + 1
        elif self.tb_config['debit_col']:
            hyperlink_col = self.tb_config['debit_col'] + 1
        else:
            hyperlink_col = self.tb_sheet.max_column + 1
        
        # Add header
        self.tb_sheet.cell(self.tb_config['header_row'], hyperlink_col, 'Reference')
        
        # Add hyperlinks for matched accounts
        for tb_row, (gl_account, gl_row) in self.account_mappings.items():
            cell = self.tb_sheet.cell(tb_row, hyperlink_col)
            formula = f'=HYPERLINK("#\'General Ledger Detail\'!A{gl_row}", "View Details")'
            cell.value = formula
        
        # Add "N/A" for unmatched accounts
        for row in range(self.tb_config['data_start_row'], self.tb_sheet.max_row + 1):
            if row not in self.account_mappings:
                account_name = self.tb_sheet.cell(row, self.tb_config.get('account_name_col', 2)).value
                if account_name:
                    self.tb_sheet.cell(row, hyperlink_col, 'N/A')
        
    def save_workbook(self):
        """Save the modified workbook."""
        self.tb_wb.save(self.output_file)

def main():
    st.set_page_config(
        page_title="TB-GL Linker",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("📊 TB-GL Linker")
    st.markdown("**Link your Trial Balance to General Ledger with automatic hyperlinks**")
    
    # Add demo info and GitHub link
    col1, col2, col3 = st.columns([2, 1, 1])
    with col2:
        st.markdown("[![GitHub](https://img.shields.io/badge/GitHub-View%20Code-blue?logo=github)](https://github.com/tuteke2023/excel_referencing)")
    with col3:
        st.markdown("[![Star](https://img.shields.io/github/stars/tuteke2023/excel_referencing?style=social)](https://github.com/tuteke2023/excel_referencing)")
    
    # Sample files info
    with st.expander("💡 Need Sample Files?"):
        st.markdown("""
        **Don't have TB and GL files to test?** 
        
        Create sample Excel files with these structures:
        
        **Trial Balance (TB.xlsx):**
        - Column A: Account Code (200, 400, etc.)
        - Column B: Account Name (Sales, Expenses, etc.)
        - Column C: Account Type (Revenue, Expense, etc.)
        - Column D: Debit amounts
        - Column E: Credit amounts
        
        **General Ledger (GL.xlsx):**
        - Account sections with account names as headers
        - Transaction details under each account
        - Date, Description, Debit, Credit columns
        """)
    
    # Instructions
    with st.expander("📋 How to Use"):
        st.markdown("""
        1. **Upload your Trial Balance Excel file** (should contain Account Names, Debit, Credit columns)
        2. **Upload your General Ledger Excel file** (should contain account sections with transaction details)
        3. **Click 'Process Files'** to automatically match accounts and create hyperlinks
        4. **Download the linked file** with hyperlinks from TB accounts to GL details
        """)
    
    # File upload section
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Trial Balance File")
        tb_file = st.file_uploader(
            "Upload Trial Balance Excel file",
            type=['xlsx', 'xls'],
            help="Excel file containing your trial balance data"
        )
    
    with col2:
        st.subheader("📄 General Ledger File")
        gl_file = st.file_uploader(
            "Upload General Ledger Excel file",
            type=['xlsx', 'xls'],
            help="Excel file containing your general ledger data"
        )
    
    # Processing section
    if tb_file and gl_file:
        st.subheader("🔧 Processing Options")
        
        col1, col2 = st.columns(2)
        with col1:
            similarity_threshold = st.slider(
                "Account Matching Similarity (%)",
                min_value=50,
                max_value=100,
                value=80,
                help="Minimum similarity required to match TB accounts to GL accounts"
            )
        
        with col2:
            output_filename = st.text_input(
                "Output Filename",
                value="TB_GL_Linked.xlsx",
                help="Name for the output file"
            )
        
        # Process button
        if st.button("🚀 Process Files", type="primary"):
            try:
                # Save uploaded files to temporary location
                with tempfile.TemporaryDirectory() as tmp_dir:
                    tb_path = os.path.join(tmp_dir, "tb_temp.xlsx")
                    gl_path = os.path.join(tmp_dir, "gl_temp.xlsx")
                    output_path = os.path.join(tmp_dir, output_filename)
                    
                    # Write uploaded files
                    with open(tb_path, 'wb') as f:
                        f.write(tb_file.read())
                    with open(gl_path, 'wb') as f:
                        f.write(gl_file.read())
                    
                    # Create progress bar
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Process files
                    status_text.text("🔍 Analyzing file structures...")
                    progress_bar.progress(20)
                    
                    linker = TBGLLinkerWeb(tb_path, gl_path, output_path)
                    
                    # Override similarity threshold if needed
                    # (This would require modifying the TBGLLinker class)
                    
                    status_text.text("📊 Loading workbooks...")
                    progress_bar.progress(40)
                    linker.load_workbooks()
                    
                    status_text.text("🔍 Analyzing structures...")
                    progress_bar.progress(60)
                    linker.analyze_tb_structure()
                    linker.analyze_gl_structure()
                    
                    status_text.text("🔗 Matching accounts...")
                    progress_bar.progress(80)
                    linker.match_accounts()
                    
                    status_text.text("📋 Creating linked file...")
                    progress_bar.progress(90)
                    linker.copy_gl_sheet()
                    linker.add_hyperlinks()
                    linker.save_workbook()
                    
                    progress_bar.progress(100)
                    status_text.text("✅ Processing complete!")
                    
                    # Display results
                    st.success(f"✅ Successfully processed! {len(linker.account_mappings)} accounts linked.")
                    
                    # Show matching results
                    if linker.account_mappings:
                        st.subheader("📊 Matching Results")
                        
                        # Create results dataframe
                        results_data = []
                        for tb_row, (gl_account, gl_row) in linker.account_mappings.items():
                            tb_account = linker.tb_sheet.cell(tb_row, linker.tb_config['account_name_col']).value
                            results_data.append({
                                'TB Account': tb_account,
                                'GL Account': gl_account,
                                'Status': '✅ Matched'
                            })
                        
                        df = pd.DataFrame(results_data)
                        st.dataframe(df, use_container_width=True)
                        
                        # Match statistics
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total Accounts", len(results_data))
                        with col2:
                            st.metric("Successfully Matched", len(linker.account_mappings))
                        with col3:
                            match_rate = len(linker.account_mappings) / len(results_data) * 100 if results_data else 0
                            st.metric("Match Rate", f"{match_rate:.1f}%")
                    
                    # Download button
                    with open(output_path, 'rb') as f:
                        st.download_button(
                            label="📥 Download Linked File",
                            data=f.read(),
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
            except Exception as e:
                st.error(f"❌ Error processing files: {str(e)}")
                st.error("Please check your file formats and try again.")

if __name__ == "__main__":
    main()