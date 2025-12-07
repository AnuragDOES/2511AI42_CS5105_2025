#!/usr/bin/env python3
"""
Launcher script for the CLI version of the Exam Scheduler.
"""

import argparse
import sys
import os

# IMPORT UPDATE: Using the new class name from the refactored allocation.py
from allocation import ExamScheduler
# IMPORT UPDATE: Using the new function name from the refactored logs.py
from logs import init_logger

def retrieve_cli_arguments():
    """Parse command line flags provided by the user."""
    cli_parser = argparse.ArgumentParser(description="Run the exam seating arrangement engine.")
    
    cli_parser.add_argument('--input', '-i', required=True, 
                            help='Path to the source Excel file containing timetable and rosters.')
    
    cli_parser.add_argument('--buffer', '-b', type=int, default=0, 
                            help='Number of empty seats to leave per room (capacity reduction).')
    
    cli_parser.add_argument('--density', '-d', choices=['Sparse', 'Dense'], default='Dense', 
                            help='Seating mode: Dense (normal) or Sparse (half capacity).')
    
    cli_parser.add_argument('--outdir', '-o', default='output', 
                            help='Target directory for generated files.')
    
    return cli_parser.parse_args()


def execute_main_process():
    cli_opts = retrieve_cli_arguments()
    
    # Initialize logging with slightly different filenames
    sys_log = init_logger('process_details.log', 'critical_errors.txt')
    sys_log.info('Initializing Seating Engine...')

    try:
        # Initialize the scheduler (Note: parameter names match the refactored allocation.py)
        scheduler_engine = ExamScheduler(
            source_path=cli_opts.input,
            gap_size=cli_opts.buffer,
            layout_mode=cli_opts.density,
            result_dir=cli_opts.outdir,
            log_handler=sys_log
        )

        # Step 1: Ingest Data
        # Call the RENAMED method from allocation.py
        scheduler_engine.load_and_parse_data()

        # Step 2: Run Allocation Logic
        # Call the RENAMED method from allocation.py
        scheduler_engine.process_all_slots()

        # Step 3: Export Excel Reports
        # Call the RENAMED method from allocation.py
        scheduler_engine.generate_excel_reports()

        # Step 4: Generate PDFs
        img_source_folder = "photos" 
        placeholder_icon = os.path.join(img_source_folder, "no_image_available.png")
        
        try:
            # Call the RENAMED method from allocation.py
            scheduler_engine.create_attendance_files(
                img_source=img_source_folder,
                default_icon=placeholder_icon
            )
        except Exception:
            sys_log.error("PDF generation encountered issues, but Excel files are likely safe.")

        sys_log.info('✅ Job finished successfully.')

    except Exception as fatal_err:
        sys_log.exception('Critical failure in main process: %s', fatal_err)
        print('The process crashed. Please check critical_errors.txt for debug info.')
        sys.exit(1)


if __name__ == '__main__':
    execute_main_process()