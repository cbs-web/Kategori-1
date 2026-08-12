import io
import urllib.error

import pytest

import jeoloji_yapay_zeka as yapay_zeka


def _http_error(code):
    return urllib.error.HTTPError(
        "https://example.invalid",
        code,
        "test",
        hdrs=None,
        fp=io.BytesIO(b'{"error":{"message":"test"}}'),
    )


def test_gemini_503_sinirli_olarak_yeniden_dener(monkeypatch):
    calls = []
    waits = []

    def fail(*_args, **_kwargs):
        calls.append(True)
        raise _http_error(503)

    monkeypatch.setattr(yapay_zeka.urllib.request, "urlopen", fail)
    monkeypatch.setattr(yapay_zeka.time, "sleep", waits.append)

    with pytest.raises(yapay_zeka.JeolojiYapayZekaHatasi, match="HTTP 503"):
        yapay_zeka._post_json(
            "https://example.invalid",
            {},
            {},
            1,
            retry_attempts=3,
            retry_http_statuses={503},
            retry_label="Gemini",
        )

    assert len(calls) == 3
    assert waits == [1, 2]


def test_gemini_400_yeniden_denenmez(monkeypatch):
    calls = []

    def fail(*_args, **_kwargs):
        calls.append(True)
        raise _http_error(400)

    monkeypatch.setattr(yapay_zeka.urllib.request, "urlopen", fail)

    with pytest.raises(yapay_zeka.JeolojiYapayZekaHatasi, match="HTTP 400"):
        yapay_zeka._post_json(
            "https://example.invalid",
            {},
            {},
            1,
            retry_attempts=3,
            retry_http_statuses={503},
            retry_label="Gemini",
        )

    assert len(calls) == 1
