#!/usr/bin/env python3
"""
Prompt Templates for Claude Code Headless Analysis

Centralized prompts for Excel structure analysis using Claude Code.
"""


class PromptTemplates:
    """Templates for Claude Code headless prompts."""

    IDENTIFY_SHEET_TYPE = """Analyze these Excel workbooks and identify which sheets contain the Trial Balance (TB) and General Ledger (GL) data.

WORKBOOK 1 (TB File):
{tb_data}

WORKBOOK 2 (GL File):
{gl_data}

Instructions:
1. Trial Balance (TB) sheets typically contain:
   - Account codes/numbers
   - Account names/descriptions
   - Debit and Credit columns with totals
   - Summary of all accounts

2. General Ledger (GL) sheets typically contain:
   - Detailed transactions
   - Account sections with headers
   - Date, Description, Debit, Credit columns
   - "Net Movement" or similar summary rows

Respond with ONLY a JSON object (no markdown, no explanation):
{{
    "tb_sheet": "exact sheet name for Trial Balance",
    "gl_sheet": "exact sheet name for General Ledger",
    "confidence": "high/medium/low",
    "reasoning": "brief explanation"
}}"""

    ANALYZE_TB_STRUCTURE = """Analyze this Trial Balance Excel sheet data and identify its structure.

SHEET DATA (CSV format with row numbers):
{data}

Instructions:
Find the following:
1. Header row - the row containing column headers like "Account", "Debit", "Credit"
2. Account column - which column contains account codes or names
3. Account name column - which column contains account descriptions (if separate from codes)
4. Debit column - which column contains debit amounts
5. Credit column - which column contains credit amounts
6. Data start row - first row of actual account data (usually header row + 1)

This may be from various accounting software (QuickBooks, Sage, Xero, NetSuite, etc.) so column names may vary:
- Account names: "Account", "Account Name", "Description", "Acct", "Account Code", "GL Account"
- Debit: "Debit", "Dr", "Debits", "Debit Amount"
- Credit: "Credit", "Cr", "Credits", "Credit Amount"

Respond with ONLY a JSON object (no markdown, no explanation):
{{
    "header_row": row_number,
    "account_col": column_number_or_null,
    "account_name_col": column_number,
    "debit_col": column_number,
    "credit_col": column_number,
    "data_start_row": row_number,
    "software_detected": "QuickBooks/Sage/Xero/NetSuite/Unknown",
    "confidence": "high/medium/low"
}}"""

    ANALYZE_GL_STRUCTURE = """Analyze this General Ledger Excel sheet data and identify its structure.

SHEET DATA (showing sample account sections):
{data}

Instructions:
Identify the GL structure:
1. How are account sections organized? (header rows, indentation, etc.)
2. Where are the column headers for Date, Description, Debit, Credit?
3. What text indicates summary/total rows? (e.g., "Net Movement", "Balance", "Total")
4. Column positions for Debit and Credit amounts

Common patterns by software:
- Xero: Account name as header, "Net Movement" row at end of section
- QuickBooks: Account headers with transactions indented
- Sage: Similar to Xero with "Movement" or "Balance" rows
- NetSuite: May use "Ending Balance" or similar

Respond with ONLY a JSON object (no markdown, no explanation):
{{
    "debit_col": column_number,
    "credit_col": column_number,
    "summary_row_text": ["Net Movement", "Balance", etc.],
    "account_header_pattern": "description of how to identify account headers",
    "software_detected": "QuickBooks/Sage/Xero/NetSuite/Unknown",
    "confidence": "high/medium/low"
}}"""

    FIND_ACCOUNT_SECTIONS = """Analyze this General Ledger data and find all account sections with their summary/net movement rows.

GL STRUCTURE INFO:
{structure_info}

SHEET DATA:
{data}

Instructions:
For each account section, find:
1. Account name (the header text)
2. Header row number
3. Summary row number (row with "Net Movement", "Balance", etc.)
4. Which column (Debit or Credit) has the non-zero summary value

Respond with ONLY a JSON object (no markdown, no explanation):
{{
    "accounts": [
        {{
            "name": "Account Name",
            "header_row": row_number,
            "summary_row": row_number,
            "summary_col": column_number,
            "summary_text": "Net Movement"
        }}
    ],
    "total_accounts_found": number
}}"""

    MATCH_ACCOUNTS = """Match Trial Balance accounts to General Ledger accounts using semantic understanding.

TRIAL BALANCE ACCOUNTS:
{tb_accounts}

GENERAL LEDGER ACCOUNTS:
{gl_accounts}

Instructions:
Match each TB account to its corresponding GL account. Consider:
1. Exact name matches
2. Similar names with slight variations (abbreviations, spacing)
3. Account codes if present
4. Accounting category context (Assets, Liabilities, Revenue, Expenses)

A match confidence of 0.8 or higher is considered a good match.

Respond with ONLY a JSON object (no markdown, no explanation):
{{
    "matches": [
        {{
            "tb_row": row_number,
            "tb_account": "TB Account Name",
            "gl_account": "GL Account Name",
            "confidence": 0.0_to_1.0,
            "match_reason": "exact/similar/code_match/semantic"
        }}
    ],
    "unmatched_tb": ["list of unmatched TB accounts"],
    "unmatched_gl": ["list of unmatched GL accounts"]
}}"""

    VERIFY_STRUCTURE = """Verify that the detected Excel structure is correct by checking sample data.

DETECTED STRUCTURE:
{structure}

SAMPLE DATA FROM DETECTED LOCATIONS:
{sample_data}

Instructions:
Verify that:
1. The header row actually contains column headers
2. The account columns contain account names/codes
3. The debit/credit columns contain numeric values
4. The data rows follow the expected pattern

Respond with ONLY a JSON object (no markdown, no explanation):
{{
    "structure_valid": true/false,
    "issues": ["list of any issues found"],
    "suggested_corrections": {{
        "field_name": "corrected_value"
    }}
}}"""
