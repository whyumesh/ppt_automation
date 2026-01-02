"""
PowerPoint Automation System
Version: 1.0.0
"""

__version__ = "1.0.0"
__author__ = "Enterprise ML Team"

from .template_analyzer import TemplateAnalyzer
from .content_parser import ContentParser
from .content_allocator import ContentAllocator
from .slide_generator import SlideGenerator
from .text_processor import TextProcessor
from .error_handler import ErrorHandler, PPTAutomationError

__all__ = [
    'TemplateAnalyzer',
    'ContentParser',
    'ContentAllocator',
    'SlideGenerator',
    'TextProcessor',
    'ErrorHandler',
    'PPTAutomationError'
]
