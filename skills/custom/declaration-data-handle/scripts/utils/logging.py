import os
import sys
import logging
from pathlib import Path
from datetime import datetime

_log_file_path = None
_debug_mode = False

def set_debug_mode(enabled):
    global _debug_mode
    _debug_mode = enabled

def get_log_path():
    global _log_file_path
    if _log_file_path is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        base_dir = os.path.dirname(script_dir)
        log_dir = os.path.join(base_dir, 'log')
        os.makedirs(log_dir, exist_ok=True)
        _log_file_path = os.path.join(log_dir, 'run.log')
    return _log_file_path

def log_message(message, level="INFO"):
    global _debug_mode
    
    try:
        import pandas as pd
        timestamp = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    except ImportError:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    output = f"[{timestamp}] [{level}] {message}"
    print(output)
    
    if _debug_mode:
        try:
            log_path = get_log_path()
            with open(log_path, 'a', encoding='utf-8') as f:
                f.write(output + '\n')
                f.flush()
        except Exception:
            pass

def setup_logging(skill_name=None):
    log_dir = Path(__file__).parent.parent / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file_path = log_dir / "run.log"

    logger = logging.getLogger(skill_name)
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter('[%(asctime)s] [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    logger.info(f"日志文件: {log_file_path.resolve()}")
    return logger