import logging
import logging.handlers

def init_logger(log_path='process_details.log', error_path='critical_errors.txt'):
    """
    Configures the system logging handlers for file and console output.
    """
    # Changed logger name from 'seating' to 'exam_system'
    app_logger = logging.getLogger('exam_system')
    app_logger.setLevel(logging.DEBUG)

    # Renamed 'fmt' to 'log_format'
    log_format = logging.Formatter('%(asctime)s | %(levelname)s | %(name)s | %(message)s')

    # Handler 1: Full Debug Log (renamed fh -> debug_handler)
    debug_handler = logging.FileHandler(log_path, mode='w')
    debug_handler.setLevel(logging.DEBUG)
    debug_handler.setFormatter(log_format)
    app_logger.addHandler(debug_handler)

    # Handler 2: Error Log (renamed eh -> err_handler)
    err_handler = logging.FileHandler(error_path, mode='w')
    err_handler.setLevel(logging.ERROR)
    err_handler.setFormatter(log_format)
    app_logger.addHandler(err_handler)

    # Handler 3: Console Output (renamed ch -> console_handler)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(log_format)
    app_logger.addHandler(console_handler)

    return app_logger