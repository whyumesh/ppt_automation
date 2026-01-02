"""
Content parsing from various input formats
"""

import pandas as pd
import json
from pathlib import Path
from typing import Dict, List, Union, Any

from .error_handler import ErrorHandler, ContentError


class ContentParser:
    """Parse structured input into presentation content"""
    
    SUPPORTED_FORMATS = ['.xlsx', '.xls', '.csv', '.json']
    
    def __init__(self, input_path: str, schema: Dict):
        self.input_path = Path(input_path)
        self.schema = schema
        self.content: Dict = {}
        self.error_handler = ErrorHandler()
        
        # Validate input file
        ErrorHandler.validate_file_exists(str(self.input_path))
        self._validate_format()
    
    def _validate_format(self):
        """Validate input file format"""
        if self.input_path.suffix not in self.SUPPORTED_FORMATS:
            raise ContentError(
                f"Unsupported format: {self.input_path.suffix}. "
                f"Supported: {', '.join(self.SUPPORTED_FORMATS)}"
            )
    
    def parse(self) -> Dict:
        """
        Parse input file based on extension
        
        Returns:
            Dictionary with structured presentation content
        """
        self.error_handler.log_info(f"Parsing content from: {self.input_path}")
        
        try:
            if self.input_path.suffix in ['.xlsx', '.xls']:
                self.content = self._parse_excel()
            elif self.input_path.suffix == '.csv':
                self.content = self._parse_csv()
            elif self.input_path.suffix == '.json':
                self.content = self._parse_json()
            
            self.error_handler.log_info(f"Parsed {len(self.content.get('slides', []))} slides")
            return self.content
            
        except Exception as e:
            self.error_handler.log_error(f"Error parsing content: {e}", e)
            raise ContentError(f"Failed to parse {self.input_path}: {e}")
    
    def _parse_excel(self) -> Dict:
        """Parse Excel file"""
        try:
            df = pd.read_excel(self.input_path)
            return self._structure_content(df)
        except Exception as e:
            raise ContentError(f"Excel parsing error: {e}")
    
    def _parse_csv(self) -> Dict:
        """Parse CSV file"""
        try:
            df = pd.read_csv(self.input_path)
            return self._structure_content(df)
        except Exception as e:
            raise ContentError(f"CSV parsing error: {e}")
    
    def _parse_json(self) -> Dict:
        """Parse JSON file"""
        try:
            with open(self.input_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Validate JSON structure
            if 'slides' not in data:
                raise ContentError("JSON must contain 'slides' array")
            
            return data
        except json.JSONDecodeError as e:
            raise ContentError(f"Invalid JSON: {e}")
        except Exception as e:
            raise ContentError(f"JSON parsing error: {e}")
    
    def _structure_content(self, df: pd.DataFrame) -> Dict:
        """
        Convert DataFrame to presentation structure
        
        Expected columns:
        - slide_type: Type of slide (title, content, section_header)
        - title: Slide title
        - content: Bullet points or content (pipe-separated or newline-separated)
        - section: Optional section name
        - notes: Optional speaker notes
        """
        # Validate required columns
        if 'title' not in df.columns:
            raise ContentError("Input must contain 'title' column")
        
        structured = {
            'presentation_title': df.iloc[0]['title'] if len(df) > 0 else 'Untitled',
            'slides': []
        }
        
        for idx, row in df.iterrows():
            slide_data = {
                'slide_number': idx + 1,
                'type': row.get('slide_type', 'content'),
                'title': str(row.get('title', '')).strip(),
                'content': self._parse_content_field(row.get('content', '')),
                'section': row.get('section', None),
                'notes': row.get('notes', '')
            }
            
            structured['slides'].append(slide_data)
        
        return structured
    
    def _parse_content_field(self, content: Any) -> List[str]:
        """
        Convert content field to list of bullet points
        
        Supports:
        - Pipe-separated: "Point 1 | Point 2 | Point 3"
        - Newline-separated: "Point 1\nPoint 2\nPoint 3"
        - Single line: "Single point"
        """
        if pd.isna(content) or content == '':
            return []
        
        content = str(content).strip()
        
        if not content:
            return []
        
        # Try pipe separator first
        if '|' in content:
            items = [item.strip() for item in content.split('|') if item.strip()]
        # Then newline
        elif '\n' in content:
            items = [item.strip() for item in content.split('\n') if item.strip()]
        # Single item
        else:
            items = [content]
        
        return items
    
    def validate(self, content: Dict = None) -> bool:
        """
        Validate content against schema
        
        Args:
            content: Content dict to validate (uses self.content if None)
            
        Returns:
            True if valid
            
        Raises:
            ContentError if validation fails
        """
        if content is None:
            content = self.content
        
        if not content:
            raise ContentError("No content to validate")
        
        # Validate required fields at top level
        if 'slides' not in content:
            raise ContentError("Content must contain 'slides' key")
        
        # Validate each slide
        required_fields = self.schema.get('required_fields', ['title'])
        
        for idx, slide in enumerate(content['slides']):
            for field in required_fields:
                if field not in slide or not slide[field]:
                    raise ContentError(
                        f"Slide {idx + 1}: Missing required field '{field}'"
                    )
            
            # Validate slide type
            valid_types = ['title', 'content', 'section_header', 'two_column', 'closing']
            slide_type = slide.get('type', 'content')
            if slide_type not in valid_types:
                self.error_handler.log_warning(
                    f"Slide {idx + 1}: Unknown slide type '{slide_type}', using 'content'"
                )
                slide['type'] = 'content'
        
        self.error_handler.log_info("Content validation passed")
        return True
    
    def get_summary(self) -> Dict:
        """Get summary statistics of parsed content"""
        if not self.content:
            return {}
        
        slides = self.content.get('slides', [])
        
        summary = {
            'total_slides': len(slides),
            'slide_types': {},
            'avg_bullets_per_slide': 0,
            'total_bullets': 0
        }
        
        total_bullets = 0
        for slide in slides:
            slide_type = slide.get('type', 'content')
            summary['slide_types'][slide_type] = summary['slide_types'].get(slide_type, 0) + 1
            
            content = slide.get('content', [])
            if isinstance(content, list):
                total_bullets += len(content)
        
        summary['total_bullets'] = total_bullets
        if len(slides) > 0:
            summary['avg_bullets_per_slide'] = round(total_bullets / len(slides), 2)
        
        return summary
