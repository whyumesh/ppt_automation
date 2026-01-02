#!/usr/bin/env python3
"""
PowerPoint Automation System - Main Entry Point

Usage:
    python main.py --template templates/corporate.pptx --input data/content.xlsx --output output/presentation.pptx
"""

import argparse
import sys
from pathlib import Path
import yaml

from src import (
    TemplateAnalyzer,
    ContentParser,
    ContentAllocator,
    SlideGenerator,
    TextProcessor,
    ErrorHandler,
    PPTAutomationError
)


def load_config(config_path: str = "config/config.yaml") -> dict:
    """Load configuration from YAML file"""
    try:
        with open(config_path, 'r') as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        print(f"⚠️ Config file not found: {config_path}, using defaults")
        return {}
    except Exception as e:
        print(f"⚠️ Error loading config: {e}, using defaults")
        return {}


def setup_directories(config: dict):
    """Create necessary directories"""
    paths = config.get('paths', {})
    
    for key, path in paths.items():
        Path(path).mkdir(parents=True, exist_ok=True)


def main():
    """Main execution function"""
    
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description='Automated PowerPoint Generation System'
    )
    parser.add_argument(
        '--template',
        type=str,
        required=True,
        help='Path to PowerPoint template (.pptx)'
    )
    parser.add_argument(
        '--input',
        type=str,
        required=True,
        help='Path to input data (Excel, CSV, or JSON)'
    )
    parser.add_argument(
        '--output',
        type=str,
        required=True,
        help='Path for output presentation (.pptx)'
    )
    parser.add_argument(
        '--config',
        type=str,
        default='config/config.yaml',
        help='Path to configuration file'
    )
    parser.add_argument(
        '--no-nlp',
        action='store_true',
        help='Disable NLP text processing'
    )
    parser.add_argument(
        '--no-cache',
        action='store_true',
        help='Disable template caching'
    )
    parser.add_argument(
        '--verbose',
        action='store_true',
        help='Enable verbose output'
    )
    
    args = parser.parse_args()
    
    # Load configuration
    config = load_config(args.config)
    
    # Setup directories
    setup_directories(config)
    
    # Initialize error handler
    error_handler = ErrorHandler(
        log_dir=config.get('paths', {}).get('logs_dir', 'logs/')
    )
    
    try:
        print("=" * 60)
        print("PowerPoint Automation System v1.0.0")
        print("=" * 60)
        
        # Step 1: Analyze Template
        print("\n🔍 Step 1: Analyzing template...")
        analyzer = TemplateAnalyzer(
            template_path=args.template,
            cache_dir=config.get('paths', {}).get('cache_dir', 'cache/')
        )
        template_structure = analyzer.analyze(use_cache=not args.no_cache)
        
        print(f"   ✓ Template analyzed: {len(template_structure['slides'])} slide layouts found")
        
        if args.verbose:
            print(f"   - Layouts: {', '.join(template_structure['layouts'])}")
        
        # Step 2: Parse Input Content
        print("\n📄 Step 2: Parsing input content...")
        parser = ContentParser(
            input_path=args.input,
            schema=config.get('content_schema', {
                'required_fields': ['title']
            })
        )
        content = parser.parse()
        parser.validate(content)
        
        content_summary = parser.get_summary()
        print(f"   ✓ Content parsed: {content_summary['total_slides']} slides")
        
        if args.verbose:
            print(f"   - Total bullets: {content_summary['total_bullets']}")
            print(f"   - Slide types: {content_summary['slide_types']}")
        
        # Step 3: Optional NLP Processing
        use_nlp = config.get('settings', {}).get('use_nlp', True) and not args.no_nlp
        
        if use_nlp:
            print("\n🧠 Step 3: Processing content with NLP...")
            text_processor = TextProcessor()
            
            processed_count = 0
            for slide in content['slides']:
                # Process long text content
                content_field = slide.get('content', [])
                
                # If content is a single long string, summarize to bullets
                if isinstance(content_field, str) and len(content_field) > 300:
                    max_bullets = config.get('settings', {}).get(
                        'max_bullets_per_slide', 6
                    )
                    slide['content'] = text_processor.summarize_to_bullets(
                        content_field,
                        max_bullets=max_bullets
                    )
                    processed_count += 1
                
                # Format existing bullets
                elif isinstance(content_field, list):
                    slide['content'] = text_processor.format_bullet_list(
                        content_field
                    )
            
            print(f"   ✓ NLP processing complete: {processed_count} slides enhanced")
        else:
            print("\n⏭️  Step 3: Skipping NLP processing (disabled)")
        
        # Step 4: Allocate Content to Slides
        print("\n📐 Step 4: Planning slide layout...")
        allocator = ContentAllocator(
            template_structure=template_structure,
            content=content,
            config=config.get('settings', {})
        )
        allocation_plan = allocator.allocate()
        
        allocation_summary = allocator.get_allocation_summary()
        print(f"   ✓ Layout planned: {allocation_summary['total_slides']} slides")
        
        if args.verbose:
            print(f"   - Layouts used: {', '.join(allocation_summary['layouts_used'])}")
        
        # Step 5: Generate Presentation
        print("\n🎨 Step 5: Generating presentation...")
        generator = SlideGenerator(template_path=args.template)
        
        # Add metadata
        generator.prs = None  # Will be loaded in generate()
        
        output_file = generator.generate(
            allocation_plan=allocation_plan,
            output_path=args.output
        )
        
        print(f"   ✓ Presentation generated successfully!")
        
        # Final Summary
        print("\n" + "=" * 60)
        print("✅ GENERATION COMPLETE")
        print("=" * 60)
        print(f"Output file: {output_file}")
        print(f"Total slides: {allocation_summary['total_slides']}")
        print(f"Input source: {args.input}")
        print(f"Template used: {args.template}")
        print("=" * 60)
        
        return 0
        
    except PPTAutomationError as e:
        error_handler.log_error(f"Automation error: {e}", e)
        print(f"\n❌ Error: {e}")
        return 1
        
    except Exception as e:
        error_handler.log_error(f"Unexpected error: {e}", e)
        print(f"\n❌ Unexpected error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
