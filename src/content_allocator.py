"""
Content allocation to slide templates
"""

from typing import List, Dict, Optional
from .error_handler import ErrorHandler, AllocationError
from .text_processor import TextProcessor


class ContentAllocator:
    """
    Allocate content to slide templates using rule-based logic
    """
    
    def __init__(
        self, 
        template_structure: Dict, 
        content: Dict,
        config: Dict = None
    ):
        self.template = template_structure
        self.content = content
        self.config = config or {}
        
        self.error_handler = ErrorHandler()
        self.text_processor = TextProcessor()
        self.allocation_plan: List[Dict] = []
        
        # Configuration
        self.max_bullets = self.config.get('max_bullets_per_slide', 6)
        self.max_bullet_length = self.config.get('max_bullet_length', 200)
    
    def allocate(self) -> List[Dict]:
        """
        Create slide-by-slide allocation plan
        
        Returns:
            List of allocation dictionaries
        """
        self.error_handler.log_info("Starting content allocation...")
        
        slides_data = self.content.get('slides', [])
        
        if not slides_data:
            raise AllocationError("No slides found in content")
        
        self.allocation_plan = []
        
        for idx, slide_data in enumerate(slides_data):
            try:
                allocation = self._allocate_single_slide(slide_data, idx)
                
                # Handle content overflow - split into multiple slides if needed
                if self._needs_split(slide_data):
                    allocations = self._split_slide_content(slide_data, idx)
                    self.allocation_plan.extend(allocations)
                else:
                    self.allocation_plan.append(allocation)
                    
            except Exception as e:
                self.error_handler.log_error(
                    f"Error allocating slide {idx + 1}: {e}", e
                )
                # Continue with next slide
                continue
        
        self.error_handler.log_info(
            f"Allocation complete: {len(self.allocation_plan)} slides planned"
        )
        
        return self.allocation_plan
    
    def _allocate_single_slide(self, slide_data: Dict, idx: int) -> Dict:
        """Allocate single slide content to template"""
        
        slide_type = slide_data.get('type', 'content')
        
        # Find matching template
        template_slide = self._find_best_template(slide_type)
        
        if not template_slide:
            self.error_handler.log_warning(
                f"No template found for type '{slide_type}', using default"
            )
            template_slide = self._get_default_template()
        
        # Map content to placeholders
        placeholder_mapping = self._map_content_to_placeholders(
            slide_data,
            template_slide
        )
        
        allocation = {
            'slide_number': idx + 1,
            'template_slide_idx': template_slide['slide_idx'],
            'layout_name': template_slide['layout_name'],
            'slide_type': slide_type,
            'content': placeholder_mapping,
            'notes': slide_data.get('notes', '')
        }
        
        return allocation
    
    def _find_best_template(self, slide_type: str) -> Optional[Dict]:
        """Find best matching template for slide type"""
        
        template_slides = self.template.get('slides', [])
        
        # Exact match
        for template in template_slides:
            if template['slide_type'] == slide_type:
                return template
        
        # Partial match (for flexibility)
        type_mappings = {
            'title': ['title', 'section_header'],
            'content': ['content', 'two_column'],
            'section_header': ['section_header', 'title'],
        }
        
        alternatives = type_mappings.get(slide_type, [])
        for alt_type in alternatives:
            for template in template_slides:
                if template['slide_type'] == alt_type:
                    return template
        
        return None
    
    def _get_default_template(self) -> Dict:
        """Get default content template"""
        template_slides = self.template.get('slides', [])
        
        # Try to find a content slide
        for template in template_slides:
            if template['slide_type'] == 'content':
                return template
        
        # Fallback to first slide with body placeholder
        for template in template_slides:
            if template.get('has_body', False):
                return template
        
        # Last resort: first slide
        if template_slides:
            return template_slides[0]
        
        raise AllocationError("No valid template found")
    
    def _map_content_to_placeholders(
        self, 
        slide_data: Dict, 
        template: Dict
    ) -> Dict:
        """Map content fields to specific placeholders"""
        
        mapping = {}
        placeholders = template.get('placeholders', [])
        
        for placeholder in placeholders:
            ph_idx = placeholder['placeholder_idx']
            ph_type = placeholder['placeholder_type']
            max_chars = placeholder['max_chars']
            
            # Handle title placeholders
            if 'TITLE' in ph_type or 'CENTERED_TITLE' in ph_type:
                title = slide_data.get('title', '')
                
                # Truncate if needed
                if len(title) > max_chars:
                    title = self.text_processor.truncate_smart(title, max_chars)
                
                mapping[ph_idx] = {
                    'text': title,
                    'type': 'title',
                    'format': 'plain'
                }
            
            # Handle body/content placeholders
            elif 'BODY' in ph_type or 'OBJECT' in ph_type:
                content_items = slide_data.get('content', [])
                
                if isinstance(content_items, str):
                    content_items = [content_items]
                
                # Fit content to placeholder
                fitted_content = self._fit_content_to_placeholder(
                    content_items,
                    max_chars
                )
                
                mapping[ph_idx] = {
                    'text': fitted_content,
                    'type': 'body',
                    'format': 'bullets'
                }
        
        return mapping
    
    def _fit_content_to_placeholder(
        self, 
        items: List[str], 
        max_chars: int
    ) -> List[str]:
        """
        Fit content items into placeholder constraints
        
        Args:
            items: List of content items
            max_chars: Maximum total characters
            
        Returns:
            Fitted list of items
        """
        fitted = []
        current_chars = 0
        
        for item in items:
            # Truncate individual items if too long
            if len(item) > self.max_bullet_length:
                item = self.text_processor.truncate_smart(
                    item, 
                    self.max_bullet_length
                )
            
            item_length = len(item)
            
            # Check if we can fit this item
            if current_chars + item_length <= max_chars:
                fitted.append(item)
                current_chars += item_length
            else:
                # Try to fit truncated version
                remaining = max_chars - current_chars
                if remaining > 50:  # Minimum viable bullet
                    truncated = self.text_processor.truncate_smart(
                        item, 
                        remaining
                    )
                    fitted.append(truncated)
                break
            
            # Limit number of bullets
            if len(fitted) >= self.max_bullets:
                break
        
        return fitted
    
    def _needs_split(self, slide_data: Dict) -> bool:
        """Check if slide content needs to be split"""
        content = slide_data.get('content', [])
        
        if isinstance(content, str):
            return False
        
        # Split if too many items
        if len(content) > self.max_bullets * 1.5:
            return True
        
        # Split if total content too long
        total_chars = sum(len(str(item)) for item in content)
        if total_chars > 2000:  # Arbitrary threshold
            return True
        
        return False
    
    def _split_slide_content(
        self, 
        slide_data: Dict, 
        idx: int
    ) -> List[Dict]:
        """Split slide content into multiple slides"""
        
        self.error_handler.log_info(
            f"Splitting slide {idx + 1} due to content overflow"
        )
        
        content = slide_data.get('content', [])
        
        if isinstance(content, str):
            content = [content]
        
        # Split content into chunks
        content_chunks = self.text_processor.split_long_content(
            content,
            max_items=self.max_bullets
        )
        
        allocations = []
        
        for chunk_idx, chunk in enumerate(content_chunks):
            # Create modified slide data for each chunk
            chunk_slide_data = slide_data.copy()
            chunk_slide_data['content'] = chunk
            
            # Modify title for continuation slides
            if chunk_idx > 0:
                original_title = slide_data.get('title', '')
                chunk_slide_data['title'] = f"{original_title} (cont.)"
            
            allocation = self._allocate_single_slide(chunk_slide_data, idx)
            allocations.append(allocation)
        
        return allocations
    
    def get_allocation_summary(self) -> Dict:
        """Get summary of allocation plan"""
        
        summary = {
            'total_slides': len(self.allocation_plan),
            'slide_types': {},
            'layouts_used': set(),
            'warnings': []
        }
        
        for allocation in self.allocation_plan:
            slide_type = allocation.get('slide_type', 'unknown')
            summary['slide_types'][slide_type] = \
                summary['slide_types'].get(slide_type, 0) + 1
            
            layout = allocation.get('layout_name')
            if layout:
                summary['layouts_used'].add(layout)
        
        summary['layouts_used'] = list(summary['layouts_used'])
        
        return summary
