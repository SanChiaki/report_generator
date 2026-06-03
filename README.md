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
