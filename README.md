# Report Generator

HTTP service for generating PPTX reports from:

- a PowerPoint template,
- an Excel-compatible component mapping JSON,
- a business payload JSON.

The PPT mapping keeps `template_id` and `component_list`. In PPT mappings, `location` is the PowerPoint selection pane shape name.

## Development

```bash
uv run --extra dev python -m pytest
```

## Run

```bash
uv run uvicorn report_generator.api:app --reload
```

LLM post-processing reads OpenAI-compatible completion settings from `.env` or the process environment:

```dotenv
API_KEY=...
BASE_URL=https://example.com/v1
MODEL=gpt-4o-mini
COMPLETION_MODE=chat
LLM_CONCURRENCY=4
```

Use `COMPLETION_MODE=completions` for legacy OpenAI `/v1/completions` compatible endpoints.
`LLM_CONCURRENCY` controls the default maximum number of concurrent LLM post-processing calls.

## Example

Build the sample template:

```bash
uv run --extra dev python examples/build_sample_template.py
```

Start the API:

```bash
uv run uvicorn report_generator.api:app --reload
```

Generate a PPTX:

```bash
curl -X POST http://127.0.0.1:8000/reports/pptx \
  -F template=@examples/sample_template.pptx \
  -F mapping=@examples/mapping.json \
  -F payload=@examples/payload.json \
  -F llmConcurrency=4 \
  --output examples/output_report.pptx
```

Create a long-running async generation task:

```bash
curl -X POST http://127.0.0.1:8000/reports/pptx/tasks \
  -F template=@examples/general_report_template.pptx \
  -F mapping=@examples/general_report_mapping.json \
  -F payload=@examples/general_report_payload.json \
  -F llmConcurrency=4
```

Then poll `/reports/pptx/tasks?taskId=...` and download from `/reports/pptx/tasks/download?taskId=...` after the task succeeds.
