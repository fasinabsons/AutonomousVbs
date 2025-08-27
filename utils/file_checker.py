#!/usr/bin/env python3
"""
MoonFlower File Checker and Recovery

Purpose
- Validate that daily CSV, Excel, and PDF outputs exist and look sane
- Provide a single CLI to run checks, generate a JSON report, and optionally clean obviously bad files
- Auto-create EHC folders if missing, so it works on any Windows PC without prior setup

Usage examples
- python utils/file_checker.py --check all
- python utils/file_checker.py --check report --out EHC_Logs/file_report.json
- python utils/file_checker.py --check cleanup
- python utils/file_checker.py --date 2024-08-08 --check pdf

Exit codes
- 0: All requested checks passed or a report was generated
- 1: One or more requested checks failed
- 2: Unexpected error
"""

from __future__ import annotations

import argparse
import json
import logging
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple


def _setup_logger() -> logging.Logger:
    logger = logging.getLogger("MoonFlowerFileChecker")
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        console = logging.StreamHandler()
        console.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(console)

        try:
            logs_dir = Path("EHC_Logs")
            logs_dir.mkdir(exist_ok=True)
            file_handler = logging.FileHandler(logs_dir / f"file_checker_{datetime.now().strftime('%Y%m%d')}.log", encoding='utf-8')
            file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            logger.addHandler(file_handler)
        except Exception:
            # Logging to file is best-effort; continue if it fails
            pass

    return logger


class PathHelper:
    """Resolve daily folders relative to project root in a portable way."""

    def __init__(self, project_root: Optional[Path] = None) -> None:
        self.project_root: Path = project_root or Path.cwd()

    def get_date_folder_name(self, target_date: Optional[str] = None) -> str:
        # ddMMM lower (e.g., 08aug)
        if target_date:
            try:
                dt = datetime.strptime(target_date, "%Y-%m-%d")
            except ValueError:
                dt = datetime.now()
        else:
            dt = datetime.now()
        return dt.strftime("%d%b").lower()

    def csv_dir(self, date_folder: str) -> Path:
        return self.project_root / "EHC_Data" / date_folder

    def excel_dir(self, date_folder: str) -> Path:
        return self.project_root / "EHC_Data_Merge" / date_folder

    def pdf_dir(self, date_folder: str) -> Path:
        return self.project_root / "EHC_Data_Pdf" / date_folder

    def ensure_daily_dirs(self, date_folder: str) -> None:
        for d in [self.csv_dir(date_folder), self.excel_dir(date_folder), self.pdf_dir(date_folder), self.project_root / "EHC_Logs"]:
            d.mkdir(parents=True, exist_ok=True)


class FileChecker:
    def __init__(self, logger: logging.Logger, path_helper: PathHelper, date_folder: str) -> None:
        self.logger = logger
        self.paths = path_helper
        self.date_folder = date_folder

        # Minimum file sizes (bytes) to catch obviously broken files
        self.minimum_sizes: Dict[str, int] = {
            "csv": 512,    # CSV should not be trivially small
            "xls": 4096,   # Old Excel format
            "xlsx": 8192,  # New Excel format
            "pdf": 10240,  # Basic sanity check
        }

    def _list_files(self, directory: Path, extensions: Tuple[str, ...]) -> List[Path]:
        if not directory.exists():
            return []
        files: List[Path] = []
        for ext in extensions:
            files.extend(directory.glob(f"*.{ext}"))
        return sorted(files)

    def _basic_file_ok(self, file_path: Path, kind: str) -> bool:
        try:
            size = file_path.stat().st_size
        except Exception:
            return False
        min_size = self.minimum_sizes.get(kind.lower(), 0)
        if size < min_size:
            self.logger.warning(f"File too small: {file_path.name} ({size} bytes < {min_size})")
            return False
        return True

    def check_csv(self) -> Dict[str, object]:
        dir_path = self.paths.csv_dir(self.date_folder)
        csv_files = self._list_files(dir_path, ("csv",))
        valid = 0
        for f in csv_files:
            if self._basic_file_ok(f, "csv"):
                valid += 1
        result = {
            "folder": str(dir_path),
            "total": len(csv_files),
            "valid_basic": valid,
            "ok": len(csv_files) >= 4 and valid >= 4,
        }
        self.logger.info(f"CSV check: {result}")
        return result

    def check_excel(self) -> Dict[str, object]:
        dir_path = self.paths.excel_dir(self.date_folder)
        excel_files = self._list_files(dir_path, ("xls", "xlsx"))
        valid = 0
        for f in excel_files:
            kind = f.suffix.lstrip('.').lower()
            if self._basic_file_ok(f, kind):
                valid += 1
        result = {
            "folder": str(dir_path),
            "total": len(excel_files),
            "valid_basic": valid,
            "ok": len(excel_files) >= 1 and valid >= 1,
        }
        self.logger.info(f"Excel check: {result}")
        return result

    def check_pdf(self) -> Dict[str, object]:
        dir_path = self.paths.pdf_dir(self.date_folder)
        pdf_files = self._list_files(dir_path, ("pdf",))
        valid = 0
        for f in pdf_files:
            if self._basic_file_ok(f, "pdf"):
                valid += 1
        result = {
            "folder": str(dir_path),
            "total": len(pdf_files),
            "valid_basic": valid,
            "ok": len(pdf_files) >= 1 and valid >= 1,
        }
        self.logger.info(f"PDF check: {result}")
        return result

    def cleanup_small_or_zero(self) -> Dict[str, int]:
        removed = 0
        checked = 0
        for dir_path, exts in [
            (self.paths.csv_dir(self.date_folder), ("csv",)),
            (self.paths.excel_dir(self.date_folder), ("xls", "xlsx")),
            (self.paths.pdf_dir(self.date_folder), ("pdf",)),
        ]:
            for f in self._list_files(dir_path, exts):
                checked += 1
                kind = f.suffix.lstrip('.').lower()
                if not self._basic_file_ok(f, kind):
                    try:
                        f.unlink(missing_ok=True)
                        removed += 1
                        self.logger.warning(f"Removed suspicious file: {f}")
                    except Exception as e:
                        self.logger.warning(f"Failed to remove {f}: {e}")
        return {"checked": checked, "removed": removed}

    def generate_report(self) -> Dict[str, object]:
        csv_info = self.check_csv()
        excel_info = self.check_excel()
        pdf_info = self.check_pdf()
        overall_ok = bool(csv_info["ok"]) and bool(excel_info["ok"]) and bool(pdf_info["ok"]) 
        return {
            "date_folder": self.date_folder,
            "csv": csv_info,
            "excel": excel_info,
            "pdf": pdf_info,
            "overall_ok": overall_ok,
            "generated_at": datetime.now().isoformat(),
        }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="MoonFlower File Checker")
    parser.add_argument("--date", type=str, default=None, help="Target date in YYYY-MM-DD (defaults to today)")
    parser.add_argument("--check", type=str, choices=["all", "csv", "excel", "pdf", "cleanup", "report"], default="all")
    parser.add_argument("--out", type=str, default=None, help="Output JSON file for --check report")
    parser.add_argument("--project-root", type=str, default=None, help="Override project root (defaults to current directory)")
    return parser.parse_args()


def main() -> int:
    logger = _setup_logger()
    try:
        args = parse_args()
        project_root = Path(args.project_root) if args.project_root else Path.cwd()
        paths = PathHelper(project_root)
        date_folder = paths.get_date_folder_name(args.date)
        paths.ensure_daily_dirs(date_folder)

        checker = FileChecker(logger, paths, date_folder)

        if args.check == "cleanup":
            stats = checker.cleanup_small_or_zero()
            logger.info(f"Cleanup completed: {stats}")
            return 0

        if args.check == "report":
            report = checker.generate_report()
            if args.out:
                try:
                    out_path = Path(args.out)
                    out_path.parent.mkdir(parents=True, exist_ok=True)
                    out_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
                    logger.info(f"Report written to: {out_path}")
                except Exception as e:
                    logger.warning(f"Failed to write report to {args.out}: {e}")
            else:
                # Print to stdout
                print(json.dumps(report, indent=2))
            return 0

        results: List[bool] = []
        if args.check in ("all", "csv"):
            results.append(bool(checker.check_csv()["ok"]))
        if args.check in ("all", "excel"):
            results.append(bool(checker.check_excel()["ok"]))
        if args.check in ("all", "pdf"):
            results.append(bool(checker.check_pdf()["ok"]))

        # If no specific result appended (shouldn't happen), treat as failure
        if not results:
            return 1

        return 0 if all(results) else 1

    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())


