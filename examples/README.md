# Example Usage

This directory contains example scripts and workflows for using the Visio Restyle application.

## Example 1: Full Conversion

This is the simplest way to convert a diagram:

```bash
python -m visio_restyle.main convert source_diagram.vsdx \
  -t target_template.vsdx \
  -o output_diagram.vsdx \
  --save-intermediate
```

This will:
1. Extract the source diagram structure
2. Extract available masters from the template
3. Use LLM to create intelligent mappings
4. Rebuild the diagram with the new style
5. Save intermediate files to `intermediate/` directory

## Example 2: Step-by-Step Conversion

For more control over the process:

```bash
# Step 1: Extract source diagram
python -m visio_restyle.main extract source_diagram.vsdx -o diagram.json

# Step 2: Extract target masters
python -m visio_restyle.main extract-masters target_template.vsdx -o masters.json

# Step 3: Generate mappings with LLM
python -m visio_restyle.main map diagram.json masters.json -o mapping.json

# Step 4: Review and edit mapping.json if needed
# Edit the file to adjust any mappings manually

# Step 5: Rebuild the diagram
python -m visio_restyle.main rebuild source_diagram.vsdx target_template.vsdx mapping.json -o output_diagram.vsdx
```

## Example 3: Using Different LLM Models

You can specify different OpenAI models:

```bash
# Use GPT-4 Turbo (faster, cheaper)
python -m visio_restyle.main convert source.vsdx -t template.vsdx -o output.vsdx -m gpt-4-turbo

# Use GPT-3.5 Turbo (fastest, cheapest, less accurate)
python -m visio_restyle.main convert source.vsdx -t template.vsdx -o output.vsdx -m gpt-3.5-turbo
```

## Example 4: Batch Processing

Create a simple script to process multiple files:

```bash
#!/bin/bash

TEMPLATE="modern_template.vsdx"

for file in *.vsdx; do
  if [ "$file" != "$TEMPLATE" ]; then
    echo "Converting $file..."
    python -m visio_restyle.main convert "$file" \
      -t "$TEMPLATE" \
      -o "converted_${file}"
  fi
done
```

## Example 5: Manual Mapping

If you want to create mappings manually without using the LLM:

```bash
# Extract what you need
python -m visio_restyle.main extract source.vsdx -o diagram.json
python -m visio_restyle.main extract-masters template.vsdx -o masters.json

# Create mapping.json manually:
# {
#   "shape_id_1": "NewMasterName1",
#   "shape_id_2": "NewMasterName2"
# }

# Rebuild
python -m visio_restyle.main rebuild source.vsdx template.vsdx mapping.json -o output.vsdx
```

## Example Files Structure

After running with `--save-intermediate`, you'll have:

```
intermediate/
├── extracted_diagram.json    # Source diagram structure
├── target_masters.json        # Available masters in template
└── shape_mapping.json         # LLM-generated mappings
```

### Example extracted_diagram.json

```json
{
  "filename": "flowchart.vsdx",
  "pages": [
    {
      "name": "Page-1",
      "shapes": [
        {
          "id": "1",
          "text": "Start Process",
          "master_name": "Process",
          "position": {"x": 4.25, "y": 8.5},
          "size": {"width": 1.5, "height": 0.75}
        }
      ],
      "connectors": [...]
    }
  ]
}
```

### Example target_masters.json

```json
{
  "masters": [
    {"name": "ModernProcess", "id": "1", "description": "..."},
    {"name": "ModernDecision", "id": "2", "description": "..."}
  ]
}
```

### Example shape_mapping.json

```json
{
  "1": "ModernProcess",
  "2": "ModernDecision",
  "3": "ModernData"
}
```

## Tips for Best Results

1. **Choose a good template**: Make sure your target template has masters that correspond to the types of shapes in your source diagram.

2. **Review mappings**: Always review the LLM-generated mappings, especially for complex diagrams.

3. **Use consistent naming**: Templates with well-named masters (e.g., "Process", "Decision") work better than generic names.

4. **Test with small diagrams first**: Start with simple diagrams to understand how the tool works.

5. **Save intermediate files**: Use `--save-intermediate` to debug and review the conversion process.

## Troubleshooting Examples

### Issue: Shape not mapped correctly

```bash
# Extract and review what the LLM sees
python -m visio_restyle.main extract source.vsdx -o diagram.json
python -m visio_restyle.main extract-masters template.vsdx -o masters.json

# Review the JSON files and manually create/edit mapping.json
# Then rebuild with your custom mapping
python -m visio_restyle.main rebuild source.vsdx template.vsdx mapping.json -o output.vsdx
```

### Issue: Master not found in template

Check what masters are available:
```bash
python -m visio_restyle.main extract-masters template.vsdx -o masters.json
cat masters.json
```

Then edit your mapping to use only masters that exist in the template.
