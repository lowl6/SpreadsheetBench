"""
Excel formula calculator using win32com (Windows only)
Forces Excel to calculate formulas and save cached values
"""
import os
import logging
from typing import Optional

try:
    from win32com.client import Dispatch
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False


def calculate_formulas(file_path: str, logger: Optional[logging.Logger] = None) -> bool:
    """
    Open Excel file, force calculation, and save.
    This ensures formula cached values are written for data_only mode.
    
    Args:
        file_path: Absolute path to Excel file
        logger: Optional logger for status messages
        
    Returns:
        True if successful, False otherwise
    """
    if not WIN32COM_AVAILABLE:
        if logger:
            logger.warning("[FORMULA CALC] win32com not available, skipping Excel calculation")
        return False
    
    if not os.path.exists(file_path):
        if logger:
            logger.error(f"[FORMULA CALC] File not found: {file_path}")
        return False
    
    try:
        if logger:
            logger.info(f"[FORMULA CALC] Opening Excel to calculate formulas in: {file_path}")
        
        # Launch Excel
        excel = Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        
        # Open workbook
        abs_path = os.path.abspath(file_path)
        wb = excel.Workbooks.Open(Filename=abs_path, UpdateLinks=False, ReadOnly=False)
        
        # Force full calculation
        wb.ForceFullCalculation = True
        excel.CalculateBeforeSave = True
        wb.Application.CalculateFull()
        
        # Save with calculated values
        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        
        if logger:
            logger.info(f"[FORMULA CALC] Successfully calculated and saved: {file_path}")
        return True
        
    except Exception as e:
        if logger:
            logger.error(f"[FORMULA CALC] Error: {e}")
        try:
            excel.Quit()
        except:
            pass
        return False
