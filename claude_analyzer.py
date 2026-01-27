#!/usr/bin/env python3
"""
Claude Code Integration for TB-GL Linker

Supports two modes:
1. Local CLI mode: Uses `claude -p` directly (when running locally)
2. API mode: Uses Claude Code API Relay service (for Streamlit Cloud)

Set CLAUDE_API_URL and CLAUDE_API_TOKEN in environment/secrets for API mode.
"""

import subprocess
import json
import re
import os
import logging
from typing import Optional

import requests

from excel_converter import ExcelToText
from prompt_templates import PromptTemplates

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ClaudeAnalyzer:
    """
    Analyzes Excel files using Claude Code.

    Automatically selects between:
    - API mode (if CLAUDE_API_URL is set)
    - CLI mode (if claude command is available locally)
    - Disabled (falls back gracefully)
    """

    def __init__(self, timeout: int = 120):
        """
        Initialize the analyzer.

        Args:
            timeout: Maximum seconds to wait for Claude response
        """
        self.timeout = timeout
        self._mode = None  # 'api', 'cli', or None
        self._api_url = os.getenv("CLAUDE_API_URL", "").strip()
        self._api_token = os.getenv("CLAUDE_API_TOKEN", "").strip()

    def _detect_mode(self) -> Optional[str]:
        """Detect which mode to use."""
        if self._mode is not None:
            return self._mode

        # Check API mode first (preferred for cloud deployments)
        if self._api_url:
            try:
                response = requests.get(
                    f"{self._api_url}/health",
                    timeout=10
                )
                if response.status_code == 200:
                    data = response.json()
                    if data.get("claude_available"):
                        self._mode = "api"
                        logger.info(f"Using Claude API mode: {self._api_url}")
                        return self._mode
            except Exception as e:
                logger.warning(f"Claude API not available: {e}")

        # Check CLI mode
        try:
            result = subprocess.run(
                ["claude", "--version"],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0:
                self._mode = "cli"
                logger.info("Using Claude CLI mode")
                return self._mode
        except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
            pass

        self._mode = None
        logger.info("Claude not available - will use fallback detection")
        return self._mode

    def is_available(self) -> bool:
        """Check if Claude is available (API or CLI)."""
        return self._detect_mode() is not None

    def _get_api_headers(self) -> dict:
        """Get headers for API requests."""
        headers = {"Content-Type": "application/json"}
        if self._api_token:
            headers["X-API-Token"] = self._api_token
        return headers

    def _run_claude_api(self, prompt: str) -> Optional[dict]:
        """Run Claude via API relay service."""
        try:
            response = requests.post(
                f"{self._api_url}/analyze",
                headers=self._get_api_headers(),
                json={"prompt": prompt, "timeout": self.timeout},
                timeout=self.timeout + 30
            )

            if response.status_code != 200:
                logger.error(f"API error: {response.status_code}")
                return None

            data = response.json()
            if data.get("success"):
                return data.get("result")
            else:
                logger.error(f"API analysis failed: {data.get('error')}")
                return None

        except Exception as e:
            logger.error(f"API request failed: {e}")
            return None

    def _run_claude_cli(self, prompt: str) -> Optional[dict]:
        """Run Claude via local CLI."""
        try:
            logger.info("Running Claude CLI analysis...")
            result = subprocess.run(
                ["claude", "-p", prompt, "--output-format", "json"],
                capture_output=True,
                text=True,
                timeout=self.timeout
            )

            if result.returncode != 0:
                logger.error(f"Claude CLI failed: {result.stderr}")
                return None

            output = json.loads(result.stdout)
            if isinstance(output, dict) and 'result' in output:
                return self._extract_json_from_result(output['result'])

            return output

        except subprocess.TimeoutExpired:
            logger.error(f"Claude CLI timed out after {self.timeout}s")
            return None
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse Claude output: {e}")
            return None
        except Exception as e:
            logger.error(f"Claude CLI error: {e}")
            return None

    def _run_claude(self, prompt: str) -> Optional[dict]:
        """Run Claude using the detected mode."""
        mode = self._detect_mode()

        if mode == "api":
            return self._run_claude_api(prompt)
        elif mode == "cli":
            return self._run_claude_cli(prompt)
        else:
            return None

    def _extract_json_from_result(self, result_text: str) -> Optional[dict]:
        """Extract JSON from Claude's result text."""
        if not result_text:
            return None

        if isinstance(result_text, dict):
            return result_text

        # Try direct JSON parse
        try:
            return json.loads(result_text)
        except json.JSONDecodeError:
            pass

        # Try to extract from markdown code blocks
        patterns = [
            r'```json\s*([\s\S]*?)\s*```',
            r'```\s*([\s\S]*?)\s*```',
            r'\{[\s\S]*\}',
        ]

        for pattern in patterns:
            match = re.search(pattern, result_text)
            if match:
                try:
                    json_str = match.group(1) if '```' in pattern else match.group(0)
                    return json.loads(json_str)
                except (json.JSONDecodeError, IndexError):
                    continue

        logger.warning(f"Could not extract JSON from result: {result_text[:200]}...")
        return None

    def identify_sheets(self, tb_wb, gl_wb) -> Optional[dict]:
        """Identify which sheets are TB vs GL."""
        tb_summary = ExcelToText.sheet_names_summary(tb_wb)
        gl_summary = ExcelToText.sheet_names_summary(gl_wb)

        tb_preview = ""
        gl_preview = ""

        if tb_wb.sheetnames:
            tb_sheet = tb_wb[tb_wb.sheetnames[0]]
            tb_preview = f"\n\nFirst sheet preview:\n{ExcelToText.sheet_to_csv_preview(tb_sheet, max_rows=20)}"

        if gl_wb.sheetnames:
            gl_sheet = gl_wb[gl_wb.sheetnames[0]]
            gl_preview = f"\n\nFirst sheet preview:\n{ExcelToText.sheet_to_csv_preview(gl_sheet, max_rows=20)}"

        prompt = PromptTemplates.IDENTIFY_SHEET_TYPE.format(
            tb_data=tb_summary + tb_preview,
            gl_data=gl_summary + gl_preview
        )

        result = self._run_claude(prompt)

        if result and 'tb_sheet' in result and 'gl_sheet' in result:
            logger.info(f"Claude identified sheets: TB='{result['tb_sheet']}', GL='{result['gl_sheet']}'")
            return result

        return None

    def analyze_tb_structure(self, tb_sheet) -> Optional[dict]:
        """Analyze Trial Balance sheet structure."""
        preview = ExcelToText.sheet_to_csv_preview(tb_sheet, max_rows=30, max_cols=12)
        prompt = PromptTemplates.ANALYZE_TB_STRUCTURE.format(data=preview)
        result = self._run_claude(prompt)

        if result and all(k in result for k in ['header_row', 'debit_col', 'credit_col']):
            logger.info(f"Claude analyzed TB structure: {result}")
            return result

        return None

    def analyze_gl_structure(self, gl_sheet) -> Optional[dict]:
        """Analyze General Ledger sheet structure."""
        preview = ExcelToText.sheet_to_csv_preview(gl_sheet, max_rows=50, max_cols=10)
        sections = ExcelToText.sample_account_sections(gl_sheet, sample_size=5)
        combined_data = f"OVERVIEW:\n{preview}\n\n{sections}"

        prompt = PromptTemplates.ANALYZE_GL_STRUCTURE.format(data=combined_data)
        result = self._run_claude(prompt)

        if result and 'debit_col' in result and 'credit_col' in result:
            logger.info(f"Claude analyzed GL structure: {result}")
            return result

        return None

    def find_account_sections(self, gl_sheet, gl_structure: dict) -> Optional[dict]:
        """Find all account sections in GL."""
        sections = ExcelToText.sample_account_sections(gl_sheet, sample_size=20)

        prompt = PromptTemplates.FIND_ACCOUNT_SECTIONS.format(
            structure_info=json.dumps(gl_structure, indent=2),
            data=sections
        )

        result = self._run_claude(prompt)

        if result and 'accounts' in result:
            logger.info(f"Claude found {len(result['accounts'])} account sections")
            return result

        return None

    def match_accounts(self, tb_accounts: list, gl_accounts: list) -> Optional[dict]:
        """Match TB accounts to GL accounts."""
        tb_formatted = "\n".join(f"Row {row}: {name}" for row, name in tb_accounts)
        gl_formatted = "\n".join(f"- {name}" for name in gl_accounts)

        prompt = PromptTemplates.MATCH_ACCOUNTS.format(
            tb_accounts=tb_formatted,
            gl_accounts=gl_formatted
        )

        result = self._run_claude(prompt)

        if result and 'matches' in result:
            logger.info(f"Claude matched {len(result['matches'])} accounts")
            return result

        return None


class ClaudeAnalyzerWithFallback:
    """Wrapper that uses Claude when available, falls back to original logic."""

    def __init__(self, timeout: int = 120):
        self.claude = ClaudeAnalyzer(timeout=timeout)
        self._use_claude = None

    def should_use_claude(self) -> bool:
        """Check if Claude should be used."""
        if self._use_claude is None:
            self._use_claude = self.claude.is_available()
        return self._use_claude

    def disable_claude(self):
        """Disable Claude analysis."""
        self._use_claude = False

    def enable_claude(self):
        """Enable Claude analysis if available."""
        self._use_claude = self.claude.is_available()
