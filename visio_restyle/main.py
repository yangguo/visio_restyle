"""
Main CLI application for Visio diagram restyling.
"""

import argparse
import sys
from pathlib import Path
from typing import Optional
import json
from dotenv import load_dotenv

from .visio_extractor import VisioExtractor
from .llm_mapper import LLMMapper
from .visio_rebuilder import VisioRebuilder


def extract_command(args):
    """Handle the extract subcommand."""
    print(f"Extracting diagram from: {args.input}")
    
    extractor = VisioExtractor(args.input)
    extractor.save_to_json(args.output)
    
    print(f"Diagram data saved to: {args.output}")


def extract_masters_command(args):
    """Handle the extract-masters subcommand."""
    print(f"Extracting masters from: {args.input}")
    
    masters = VisioExtractor.get_masters_from_visio(args.input)
    
    with open(args.output, 'w', encoding='utf-8') as f:
        json.dump({"masters": masters}, f, indent=2, ensure_ascii=False)
    
    print(f"Found {len(masters)} masters")
    print(f"Masters saved to: {args.output}")


def map_command(args):
    """Handle the map subcommand."""
    print(f"Creating shape mapping using LLM...")
    
    # Load diagram data
    with open(args.diagram, 'r', encoding='utf-8') as f:
        diagram_data = json.load(f)
    
    # Load target masters
    with open(args.masters, 'r', encoding='utf-8') as f:
        masters_data = json.load(f)
        target_masters = masters_data.get("masters", [])
    
    # Create mapper
    mapper = LLMMapper(model=args.model)
    
    # Generate mapping
    mapping = mapper.create_mapping(diagram_data, target_masters)
    
    # Save mapping
    mapper.save_mapping(mapping, args.output)
    
    print(f"Mapping created with {len(mapping)} shape mappings")
    print(f"Mapping saved to: {args.output}")


def rebuild_command(args):
    """Handle the rebuild subcommand."""
    print(f"Rebuilding diagram with new masters...")
    
    # Load mapping
    mapping = LLMMapper.load_mapping(args.mapping)
    
    # Create rebuilder
    rebuilder = VisioRebuilder(
        source_path=args.input,
        target_template_path=args.template,
        mapping=mapping
    )
    
    # Rebuild diagram
    rebuilder.rebuild(args.output)
    
    print(f"Rebuilt diagram saved to: {args.output}")


def convert_command(args):
    """Handle the convert subcommand (full workflow)."""
    print("=" * 60)
    print("VISIO RESTYLE - Full Conversion Workflow")
    print("=" * 60)
    
    # Step 1: Extract source diagram
    print("\n[1/4] Extracting source diagram...")
    extractor = VisioExtractor(args.input)
    diagram_data = extractor.extract()
    
    # Step 2: Extract target masters
    print(f"[2/4] Extracting masters from template: {args.template}")
    target_masters = VisioExtractor.get_masters_from_visio(args.template)
    print(f"      Found {len(target_masters)} available masters")
    
    # Step 3: Create mapping with LLM
    print("[3/4] Generating shape mappings with LLM...")
    mapper = LLMMapper(model=args.model)
    mapping = mapper.create_mapping(diagram_data, target_masters)
    print(f"      Created {len(mapping)} shape mappings")
    
    # Optionally save intermediate files
    if args.save_intermediate:
        intermediate_dir = Path(args.output).parent / "intermediate"
        intermediate_dir.mkdir(exist_ok=True)
        
        diagram_path = intermediate_dir / "extracted_diagram.json"
        masters_path = intermediate_dir / "target_masters.json"
        mapping_path = intermediate_dir / "shape_mapping.json"
        
        with open(diagram_path, 'w', encoding='utf-8') as f:
            json.dump(diagram_data, f, indent=2, ensure_ascii=False)
        
        with open(masters_path, 'w', encoding='utf-8') as f:
            json.dump({"masters": target_masters}, f, indent=2, ensure_ascii=False)
        
        mapper.save_mapping(mapping, str(mapping_path))
        
        print(f"      Intermediate files saved to: {intermediate_dir}")
    
    # Step 4: Rebuild diagram
    print("[4/4] Rebuilding diagram with new style...")
    rebuilder = VisioRebuilder(
        source_path=args.input,
        target_template_path=args.template,
        mapping=mapping
    )
    rebuilder.rebuild(args.output)
    
    print("\n" + "=" * 60)
    print(f"✓ Conversion complete!")
    print(f"✓ Output saved to: {args.output}")
    print("=" * 60)


def main():
    """Main entry point for the CLI."""
    # Load environment variables from .env file
    load_dotenv()
    
    parser = argparse.ArgumentParser(
        description="Convert Visio diagrams to different visual styles using LLM-powered shape mapping",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Full conversion workflow
  %(prog)s convert input.vsdx -t template.vsdx -o output.vsdx
  
  # Step-by-step workflow
  %(prog)s extract input.vsdx -o diagram.json
  %(prog)s extract-masters template.vsdx -o masters.json
  %(prog)s map diagram.json masters.json -o mapping.json
  %(prog)s rebuild input.vsdx template.vsdx mapping.json -o output.vsdx
        """
    )
    
    subparsers = parser.add_subparsers(dest='command', help='Available commands')
    
    # Extract command
    extract_parser = subparsers.add_parser(
        'extract',
        help='Extract diagram structure to JSON'
    )
    extract_parser.add_argument('input', help='Input .vsdx file')
    extract_parser.add_argument('-o', '--output', required=True, help='Output JSON file')
    extract_parser.set_defaults(func=extract_command)
    
    # Extract masters command
    extract_masters_parser = subparsers.add_parser(
        'extract-masters',
        help='Extract available masters from a Visio file'
    )
    extract_masters_parser.add_argument('input', help='Input .vsdx file')
    extract_masters_parser.add_argument('-o', '--output', required=True, help='Output JSON file')
    extract_masters_parser.set_defaults(func=extract_masters_command)
    
    # Map command
    map_parser = subparsers.add_parser(
        'map',
        help='Create shape mapping using LLM'
    )
    map_parser.add_argument('diagram', help='Extracted diagram JSON file')
    map_parser.add_argument('masters', help='Target masters JSON file')
    map_parser.add_argument('-o', '--output', required=True, help='Output mapping JSON file')
    map_parser.add_argument('-m', '--model', default='gpt-4', help='LLM model to use (default: gpt-4)')
    map_parser.set_defaults(func=map_command)
    
    # Rebuild command
    rebuild_parser = subparsers.add_parser(
        'rebuild',
        help='Rebuild diagram with new masters'
    )
    rebuild_parser.add_argument('input', help='Source .vsdx file')
    rebuild_parser.add_argument('template', help='Target template .vsdx file')
    rebuild_parser.add_argument('mapping', help='Shape mapping JSON file')
    rebuild_parser.add_argument('-o', '--output', required=True, help='Output .vsdx file')
    rebuild_parser.set_defaults(func=rebuild_command)
    
    # Convert command (full workflow)
    convert_parser = subparsers.add_parser(
        'convert',
        help='Full conversion workflow (extract -> map -> rebuild)'
    )
    convert_parser.add_argument('input', help='Source .vsdx file')
    convert_parser.add_argument('-t', '--template', required=True, help='Target template .vsdx file')
    convert_parser.add_argument('-o', '--output', required=True, help='Output .vsdx file')
    convert_parser.add_argument('-m', '--model', default='gpt-4', help='LLM model to use (default: gpt-4)')
    convert_parser.add_argument(
        '--save-intermediate',
        action='store_true',
        help='Save intermediate JSON files (diagram, masters, mapping)'
    )
    convert_parser.set_defaults(func=convert_command)
    
    # Parse arguments
    args = parser.parse_args()
    
    # Execute command
    if not args.command:
        parser.print_help()
        sys.exit(1)
    
    try:
        args.func(args)
    except Exception as e:
        print(f"\nError: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
