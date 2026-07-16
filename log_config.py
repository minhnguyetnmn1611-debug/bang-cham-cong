import logging
import os
from logging.handlers import RotatingFileHandler

def get_logger(name="vmos_app"):
    logger = logging.getLogger(name)
    
    # Only configure if it doesn't already have handlers to avoid duplication
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        
        # Console Handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        logger.addHandler(ch)
        
        # File Handler (Rotating, max 5MB, keep 3 backups)
        try:
            fh = RotatingFileHandler(
                "app.log", maxBytes=5*1024*1024, backupCount=3, encoding="utf-8"
            )
            fh.setLevel(logging.WARNING) # Store warnings and errors in file
            fh.setFormatter(formatter)
            logger.addHandler(fh)
        except Exception as e:
            # Fallback to console if file writing fails
            print(f"Cannot setup file logger: {e}")
            
    return logger

logger = get_logger()
