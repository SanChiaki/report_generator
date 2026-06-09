import httpx
import pytest

from report_generator.errors import ErrorCode, ReportGenerationError
from report_generator.llm import OpenAICompletionProcessor
from report_generator.models import ComponentMapping


def text_component():
    return ComponentMapping.model_validate(
        {
            "location": "text.summary",
            "semantic_description": "项目摘要",
            "prompt": "输出一句总结",
            "data_example": "项目正常推进",
            "type": "Text",
        }
    )


def test_openai_completion_processor_posts_component_context(monkeypatch):
    calls = []

    def fake_post(url, json, headers, timeout):
        calls.append({"url": url, "json": json, "headers": headers, "timeout": timeout})
        return httpx.Response(
            200,
            json={"choices": [{"message": {"content": '{"data":"项目正常推进"}'}}]},
            request=httpx.Request("POST", url),
        )

    monkeypatch.setattr(httpx, "post", fake_post)
    processor = OpenAICompletionProcessor(
        api_key="test-key",
        base_url="https://llm.example.com/v1",
        model="test-model",
    )

    result = processor.process(text_component(), {"status": "green"})

    assert result == "项目正常推进"
    assert calls[0]["url"] == "https://llm.example.com/v1/chat/completions"
    assert calls[0]["headers"]["Authorization"] == "Bearer test-key"
    assert calls[0]["json"]["model"] == "test-model"
    user_content = calls[0]["json"]["messages"][1]["content"]
    assert "项目摘要" in user_content
    assert "输出一句总结" in user_content
    assert "项目正常推进" in user_content
    assert "green" in user_content


def test_openai_completion_processor_requires_data_json(monkeypatch):
    monkeypatch.setattr(
        httpx,
        "post",
        lambda *args, **kwargs: httpx.Response(
            200,
            json={"choices": [{"message": {"content": '{"value":"missing data"}'}}]},
            request=httpx.Request("POST", args[0]),
        ),
    )
    processor = OpenAICompletionProcessor(
        api_key="test-key",
        base_url="https://llm.example.com/v1",
        model="test-model",
    )

    with pytest.raises(ReportGenerationError) as exc:
        processor.process(text_component(), {"status": "green"})

    assert exc.value.error_code == ErrorCode.POST_PROCESSING_FAILED


def test_openai_completion_processor_reads_dotenv(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    tmp_path.joinpath(".env").write_text(
        "API_KEY=dotenv-key\nBASE_URL=https://dotenv.example.com/v1\nMODEL=dotenv-model\n",
        encoding="utf-8",
    )

    processor = OpenAICompletionProcessor()

    assert processor.api_key == "dotenv-key"
    assert processor.base_url == "https://dotenv.example.com/v1"
    assert processor.model == "dotenv-model"


def test_openai_completion_processor_supports_legacy_completions(monkeypatch):
    calls = []

    def fake_post(url, json, headers, timeout):
        calls.append({"url": url, "json": json, "headers": headers, "timeout": timeout})
        return httpx.Response(
            200,
            json={"choices": [{"text": '{"data":"项目正常推进"}'}]},
            request=httpx.Request("POST", url),
        )

    monkeypatch.setattr(httpx, "post", fake_post)
    processor = OpenAICompletionProcessor(
        api_key="test-key",
        base_url="https://llm.example.com/v1",
        model="test-model",
        completion_mode="completions",
    )

    result = processor.process(text_component(), {"status": "green"})

    assert result == "项目正常推进"
    assert calls[0]["url"] == "https://llm.example.com/v1/completions"
    assert calls[0]["json"]["prompt"]
