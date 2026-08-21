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


def test_gemini_400_ayrintiyi_kullaniciya_aktarir(monkeypatch):
    def fail(*_args, **_kwargs):
        raise urllib.error.HTTPError(
            "https://example.invalid",
            400,
            "test",
            hdrs=None,
            fp=io.BytesIO(
                b'{"error":{"message":"Unknown field",'
                b'"details":[{"fieldViolations":[{"field":"generationConfig.responseJsonSchema",'
                b'"description":"unsupported"}]}]}}'
            ),
        )

    monkeypatch.setattr(yapay_zeka.urllib.request, "urlopen", fail)

    with pytest.raises(
        yapay_zeka.JeolojiYapayZekaHatasi,
        match=r"generationConfig\.responseJsonSchema: unsupported",
    ):
        yapay_zeka._post_json(
            "https://example.invalid",
            {},
            {},
            1,
            retry_label="Gemini",
        )


def test_gemini_36_rest_yapilandirilmis_cikti_bicimini_kullanir(monkeypatch):
    captured = {}

    def fake_post(_url, _headers, payload, _timeout, **_kwargs):
        captured.update(payload)
        return {
            "candidates": [
                {
                    "content": {
                        "parts": [
                            {
                                "text": (
                                    '{"birimler":[],"ana_parsel_kodu":"",'
                                    '"genel_guven":0,"notlar":""}'
                                )
                            }
                        ]
                    }
                }
            ]
        }

    monkeypatch.setattr(yapay_zeka, "_post_json", fake_post)

    yapay_zeka._gemini_cagir(
        {
            "katalog": [{"kod": "Kça", "ad": "Çamlıca Metamorfitleri"}],
            "gorseller": [],
            "yerel_birimler": [],
            "pafta_adlari": [],
        },
        "test-key",
        "gemini-3.6-flash",
        1,
    )

    config = captured["generationConfig"]
    assert config["responseFormat"]["text"]["mimeType"] == "APPLICATION_JSON"
    assert "schema" not in config["responseFormat"]["text"]
    assert "responseMimeType" not in config
    assert "responseJsonSchema" not in config


def test_gemini_prompt_json_cikis_sozlesmesini_icerir():
    prompt = yapay_zeka._prompt(
        {
            "katalog": [{"kod": "Kça", "ad": "Çamlıca Metamorfitleri"}],
            "yerel_birimler": [],
            "pafta_adlari": [],
        }
    )

    assert "Yalnız geçerli bir JSON nesnesi döndür" in prompt
    assert '"birimler"' in prompt
    assert '"ana_parsel_kodu"' in prompt
