"""
Error handling and logging utilities
"""

import logging
from typing import Optional
from pathlib import Path


class PPTAutomationError(Exception):
    """Base exception for PPT automation system"""
    pass


class TemplateError(PPTAutomationError):
    """Template-related errors"""
    pass


class ContentError(PPTAutomationError):
    """Content parsing/validation errors"""
    pass


class AllocationError(PPTAutomationError):
    """Content allocation errors"""
    pass


class ErrorHandler:
    """Centralized error handling and logging"""
    
    def __init__(self, log_dir: str = "logs/"):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        self._setup_logging()
    
    def _setup_logging(self):
        """Configure logging"""
        log_file = self.log_dir / "ppt_automation.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )
        
        self.logger = logging.getLogger(__name__)
    
    def log_info(self, message: str):
        """Log info message"""
        self.logger.info(message)
    
    def log_warning(self, message: str):
        """Log warning"""
        self.logger.warning(message)
    
    def log_error(self, message: str, exc_info: Optional[Exception] = None):
        """Log error"""
        self.logger.error(message, exc_info=exc_info)
    
    @staticmethod
    def handle_missing_placeholder(slide_idx: int, ph_idx: int) -> dict:
        """Handle missing placeholder"""
        msg = f"⚠️ Slide {slide_idx}: Placeholder {ph_idx} missing"
        logging.warning(msg)
        return {
            'status': 'warning',
            'message': msg,
            'fallback': 'text_box'
        }
    
    @staticmethod
    def handle_content_overflow(content_length: int, max_chars: int) -> dict:
        """Handle content overflow"""
        msg = f"⚠️ Content truncated: {content_length} chars → {max_chars} chars"
        logging.warning(msg)
        return {
            'status': 'truncated',
            'message': msg,
            'original_length': content_length,
            'truncated_length': max_chars
        }
    
    @staticmethod
    def validate_file_exists(file_path: str) -> bool:
        """Validate file exists"""
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")
        return True
    
    @staticmethod
    def validate_template_structure(template_data: dict) -> bool:
        """Validate template structure"""
        required_keys = ['slides', 'theme']
        
        for key in required_keys:
            if key not in template_data:
                raise TemplateError(f"Invalid template: missing '{key}' section")
        
        if not template_data['slides']:
            raise TemplateError("Template has no slides")
        
        return True
