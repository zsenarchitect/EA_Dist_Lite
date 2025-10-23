#!/usr/bin/python
# -*- coding: utf-8 -*-

"""
Excel Writeback Module - Handles writing Revit area data back to Excel
"""

import os
import shutil
import time
import config
from EnneadTab import EXCEL, DATA_FILE, EXE


def fake_write_design_values(excel_data, all_matches, excel_path, worksheet):
    """
    Print (fake write) what values would be written to the DESIGN column in Excel.
    Does not actually modify the Excel file - only prints the planned updates.
    
    Args:
        excel_data (dict): Dictionary of Excel data with RowData objects
        all_matches (dict): Dictionary of matches by scheme {scheme_name: {'matches': [...]}}
        excel_path (str): Path to Excel file
        worksheet (str): Worksheet name
        
    Returns:
        dict: Statistics about the fake write operation
    """
    print("\n" + "=" * 80)
    print("=== FAKE EXCEL WRITEBACK (No actual changes) ===")
    print("=" * 80)
    
    # Read Excel to find DESIGN column index
    raw_data = EXCEL.read_data_from_excel(
        excel_path, 
        worksheet=worksheet, 
        return_dict=True
    )
    
    header_map = EXCEL.get_header_map(raw_data, config.EXCEL_HEADER_ROW)
    
    # Find DESIGN column index
    design_col = None
    for col, header in header_map.items():
        if header == "DESIGN":
            design_col = col
            break
    
    if design_col is None:
        print("ERROR: Could not find 'DESIGN' column in Excel file")
        print("Available columns: {}".format(", ".join(header_map.values())))
        return {
            'total_updates': 0,
            'error': 'DESIGN column not found'
        }
    
    # Convert column index to letter
    design_col_letter = EXCEL.column_number_to_letter(design_col)
    
    print("Found DESIGN column at index {} (Column {})".format(design_col, design_col_letter))
    print("")
    
    # Track statistics
    total_updates = 0
    updates_by_scheme = {}
    
    # Loop through all schemes and their matches
    for scheme_name, scheme_data in all_matches.items():
        matches = scheme_data.get('matches', [])
        
        if not matches:
            continue
            
        updates_by_scheme[scheme_name] = 0
        print("--- Scheme: {} ---".format(scheme_name))
        
        for match in matches:
            excel_row = match.get('excel_row_index')
            actual_dgsf = match.get('actual_dgsf', 0)
            room_name = match.get('room_name', 'Unknown')
            department = match.get('department', 'Unknown')
            division = match.get('division', 'Unknown')
            
            # Print the fake write information
            print("\n\n[{}] Row {}, Column DESIGN ({}): {:.2f}".format(
                scheme_name,
                excel_row,
                design_col_letter,
                float(actual_dgsf)
            ))
            print("  Entry: {} | {} | {}".format(
                department,
                division,
                room_name
            ))
            print("")
            
            total_updates += 1
            updates_by_scheme[scheme_name] += 1
    
    # Print summary
    print("=" * 80)
    print("SUMMARY:")
    print("  Total cells that would be updated: {}".format(total_updates))
    for scheme_name, count in updates_by_scheme.items():
        print("    {}: {} cells".format(scheme_name, count))
    print("=" * 80)
    print("")
    
    return {
        'total_updates': total_updates,
        'updates_by_scheme': updates_by_scheme,
        'design_column': design_col_letter,
        'error': None
    }


def write_design_values_to_excel(excel_data, all_matches, excel_path, worksheet):
    """
    Write actual Revit DGSF values to the DESIGN column in the Excel file.
    Writes directly to the original Excel file.
    
    Args:
        excel_data (dict): Dictionary of Excel data with RowData objects
        all_matches (dict): Dictionary of matches by scheme {scheme_name: {'matches': [...]}}
        excel_path (str): Path to Excel file
        worksheet (str): Worksheet name
        
    Returns:
        dict: Statistics about the write operation
    """
    print("\n" + "=" * 80)
    print("=== REAL EXCEL WRITEBACK TO ORIGINAL FILE ===")
    print("=" * 80)
    print("Target file: {}".format(excel_path))
    print("")
    
    # Read Excel to find DESIGN column index
    raw_data = EXCEL.read_data_from_excel(
        excel_path, 
        worksheet=worksheet, 
        return_dict=True
    )
    
    header_map = EXCEL.get_header_map(raw_data, config.EXCEL_HEADER_ROW)
    
    # Find DESIGN column index
    design_col = None
    for col, header in header_map.items():
        if header == "DESIGN":
            design_col = col
            break
    
    if design_col is None:
        error_msg = "DESIGN column not found"
        print("ERROR: Could not find 'DESIGN' column in Excel file")
        print("Available columns: {}".format(", ".join(header_map.values())))
        return {
            'total_updates': 0,
            'error': error_msg
        }
    
    # Convert column index to letter
    design_col_letter = EXCEL.column_number_to_letter(design_col)
    
    print("Found DESIGN column at index {} (Column {})".format(design_col, design_col_letter))
    print("")
    
    # Build update_data dictionary for ExcelHandler
    update_data_dict = {}
    total_updates = 0
    updates_by_scheme = {}
    
    # Loop through all schemes and their matches
    for scheme_name, scheme_data in all_matches.items():
        matches = scheme_data.get('matches', [])
        
        if not matches:
            continue
            
        updates_by_scheme[scheme_name] = 0
        print("--- Processing Scheme: {} ---".format(scheme_name))
        
        for match in matches:
            excel_row = match.get('excel_row_index')
            actual_dgsf = match.get('actual_dgsf', 0)
            room_name = match.get('room_name', 'Unknown')
            department = match.get('department', 'Unknown')
            division = match.get('division', 'Unknown')
            
            # Create update entry for ExcelHandler
            cell_key = "{},{}".format(excel_row, design_col_letter)
            update_data_dict[cell_key] = {
                "value": float(actual_dgsf),
                "row": excel_row,
                "column": design_col
            }
            
            # Print what we're writing
            print("Writing to Row {}, Column DESIGN ({}): {:.2f}".format(
                excel_row,
                design_col_letter,
                float(actual_dgsf)
            ))
            print("  Entry: {} | {} | {}".format(
                department,
                division,
                room_name
            ))
            print("")
            
            total_updates += 1
            updates_by_scheme[scheme_name] += 1
    
    # Prepare job data for ExcelHandler
    job_data = {
        "mode": "update",
        "filepath": excel_path,
        "worksheet": worksheet,
        "data": {
            "update_data": update_data_dict, 
            "append_data": {}
        }
    }
    
    print("=" * 80)
    print("Sending {} cell updates to ExcelHandler...".format(total_updates))
    DATA_FILE.set_data(job_data, "excel_handler_input")
    
    # Launch ExcelHandler
    EXE.try_open_app("ExcelHandler")
    
    # Wait for ExcelHandler to complete (max 30 seconds)
    max_wait = 300  # 30 seconds (300 * 0.1)
    wait = 0
    while wait < max_wait:
        job_status = DATA_FILE.get_data("excel_handler_input")
        if job_status and job_status.get("status") == "done":
            print("ExcelHandler completed successfully!")
            break
        time.sleep(0.1)
        wait += 1
    
    if wait >= max_wait:
        print("WARNING: ExcelHandler timeout after 30 seconds")
    
    # Print summary
    print("=" * 80)
    print("SUMMARY:")
    print("  Total cells updated: {}".format(total_updates))
    for scheme_name, count in updates_by_scheme.items():
        print("    {}: {} cells".format(scheme_name, count))
    print("  Original file: {}".format(excel_path))
    print("=" * 80)
    print("")
    
    # Open the Excel file
    if os.path.exists(excel_path):
        print("Opening Excel file...")
        os.startfile(excel_path)
    else:
        print("WARNING: Excel file not found at: {}".format(excel_path))
    
    return {
        'total_updates': total_updates,
        'updates_by_scheme': updates_by_scheme,
        'design_column': design_col_letter,
        'error': None
    }

