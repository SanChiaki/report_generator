# Report Generator

HTTP service for generating PPTX reports from:

- a PowerPoint template,
- an Excel-compatible component mapping JSON,
- a business payload JSON.

The PPT mapping keeps `template_id` and `component_list`. In PPT mappings, `location` is the PowerPoint selection pane shape name.

## Development

```bash
python -m pip install -e ".[dev]"
python -m pytest
```

## Run

```bash
uvicorn report_generator.api:app --reload
```

## Example

Build the sample template:

```bash
python examples/build_sample_template.py
```

Start the API:

```bash
uvicorn report_generator.api:app --reload
```

Generate a PPTX:

```bash
curl -X POST http://127.0.0.1:8000/reports/pptx \
  -F template=@examples/sample_template.pptx \
  -F mapping=@examples/mapping.json \
  -F payload=@examples/payload.json \
  --output examples/output_report.pptx
```
