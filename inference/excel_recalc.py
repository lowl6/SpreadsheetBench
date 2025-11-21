"""Excel recalculation helper for dynamic array formulas (Plan 1).

Opens a workbook in real Excel via COM (Windows only), triggers full
recalculation, and optionally materializes dynamic array spill ranges
into static values for better compatibility with libraries that don't
fully parse spill metadata (e.g. openpyxl < full dynamic array support).

Usage (programmatic):
    from excel_recalc import recalc_workbook
    recalc_workbook(path_in, materialize_dynamic=True)

CLI:
    python -m inference.excel_recalc --input your.xlsx --materialize-dynamic

Flags:
    --materialize-dynamic  If set, spilled ranges are converted to plain
                            static values (original formula cell retains
                            its formula unless --strip-formula specified).
    --strip-formula        If used together with materialization, the
                            source dynamic array formula cell is replaced
                            by its calculated top-left value to avoid any
                            future spill changes.

Notes:
    - Requires pywin32; silently no-ops if unavailable.
    - Safe to call when no dynamic arrays exist; cost is just a single
      Excel open/close cycle.
    - Dynamic spill detection uses HasSpill/SpillRange properties
      (available in Microsoft 365 / newer Excel). If not present, the
      materialization step is skipped.
"""
from __future__ import annotations

import os
import argparse
import logging
from typing import Optional

try:
    import pythoncom  # type: ignore
    from win32com.client import DispatchEx  # type: ignore
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False


def _init_excel(visible: bool = False):
    excel = DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    return excel


def _materialize_spills(wb, logger: Optional[logging.Logger], strip_formula: bool):
    """Enumerate sheets and materialize dynamic array spill ranges.

    For each cell with HasSpill=True, copy values from SpillRange into
    itself and underlying cells so that openpyxl will later see static
    values. Optionally replace the source formula with its evaluated
    value (strip_formula).
    """
    total_spills = 0
    for ws in wb.Worksheets:
        used = ws.UsedRange
        # Iterate cells; COM collections are 1-based
        for cell in used.Cells:
            try:
                # Some Excel versions may not expose HasSpill
                has_spill = getattr(cell, "HasSpill", False)
            except Exception:
                has_spill = False
            if not has_spill:
                continue
            try:
                spill_range = cell.SpillRange  # top-left formula cell's range
            except Exception:
                continue
            total_spills += 1
            # Copy values so they become static
            for sc in spill_range.Cells:
                sc.Value = sc.Value  # Assign value to itself to detach
            if strip_formula:
                # Replace formula cell with its concrete value
                try:
                    cell.Value = cell.Value
                    # Clear formula explicitly if property exists
                    cell.Formula = cell.Value
                except Exception:
                    pass
    if logger:
        logger.info(f"[EXCEL RECALC] Materialized {total_spills} dynamic spill(s)")
    return total_spills


def recalc_workbook(
    input_path: str,
    output_path: Optional[str] = None,
    visible: bool = False,
    materialize_dynamic: bool = False,
    strip_formula: bool = False,
    logger: Optional[logging.Logger] = None,
    timeout_sec: int = 60,
) -> bool:
    """Open workbook in Excel, force full calc, optional dynamic spill materialization.

    Returns True if successful.
    """
    if not WIN32_AVAILABLE:
        if logger:
            logger.warning("[EXCEL RECALC] pywin32 not available; skipping")
        return False
    if not os.path.exists(input_path):
        if logger:
            logger.error(f"[EXCEL RECALC] File not found: {input_path}")
        return False

    # COM init (explicit for multi-thread safety)
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass

    excel = None
    try:
        if logger:
            logger.info(f"[EXCEL RECALC] Opening Excel for: {input_path}")
        excel = _init_excel(visible=visible)
        abs_in = os.path.abspath(input_path)
        wb = excel.Workbooks.Open(Filename=abs_in, UpdateLinks=False, ReadOnly=False)

        wb.ForceFullCalculation = True
        excel.CalculateBeforeSave = True
        excel.CalculationInterruptKey = 0  # Disable user interrupt
        excel.CalculateFull()

        # Wait calculation state
        import time
        start = time.time()
        while getattr(excel, "CalculationState", 0) != 0 and time.time() - start < timeout_sec:
            time.sleep(0.2)

        if materialize_dynamic:
            _materialize_spills(wb, logger, strip_formula=strip_formula)

        target = output_path or input_path
        wb.SaveAs(os.path.abspath(target))
        wb.Close(SaveChanges=True)
        excel.Quit()
        if logger:
            logger.info(f"[EXCEL RECALC] Recalculation complete: {target}")
        return True
    except Exception as e:
        if logger:
            logger.error(f"[EXCEL RECALC] Error: {e}")
        try:
            if excel:
                excel.Quit()
        except Exception:
            pass
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _build_arg_parser():
    p = argparse.ArgumentParser("Excel recalculation helper")
    p.add_argument("--input", required=True, help="Input workbook path")
    p.add_argument("--output", help="Optional output path (defaults overwrite input)")
    p.add_argument("--visible", action="store_true", help="Show Excel window while recalculating")
    p.add_argument("--materialize-dynamic", action="store_true", help="Convert dynamic spill ranges to static values")
    p.add_argument("--strip-formula", action="store_true", help="Replace source dynamic formula cell with value when materializing")
    return p


def main():  # CLI entry
    parser = _build_arg_parser()
    opt = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="[%(levelname)s] %(message)s")
    logger = logging.getLogger("excel_recalc")
    recalc_workbook(
        input_path=opt.input,
        output_path=opt.output,
        visible=opt.visible,
        materialize_dynamic=opt.materialize_dynamic,
        strip_formula=opt.strip_formula,
        logger=logger,
    )


if __name__ == "__main__":
    main()
