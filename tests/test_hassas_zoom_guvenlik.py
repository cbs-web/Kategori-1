import io
import threading
from collections import OrderedDict

from PIL import Image
import requests

import harita_hassas_zoom as hassas_zoom
from harita_hassas_zoom import (
    RAW_TILE_CACHE_LIMIT,
    SCALED_PHOTO_CACHE_LIMIT,
    KesirliZoomHaritaView,
)


def _bare_widget():
    widget = object.__new__(KesirliZoomHaritaView)
    widget._fractional_main_thread_id = threading.get_ident()
    widget._fractional_lock = threading.RLock()
    widget._fractional_raw_tile_cache = OrderedDict()
    widget._fractional_photo_cache = OrderedDict()
    widget._fractional_server_generation = 0
    widget._fractional_render_generation = 1
    widget._fractional_empty_raw_image = Image.new("RGB", (256, 256), (190, 190, 190))
    widget.tile_server = "old/{z}/{x}/{y}.png"
    widget.overlay_tile_server = None
    widget.tile_size = 256
    widget.max_zoom = 22
    widget.zoom = 15.25
    widget.use_database_only = False
    widget.running = True
    widget.image_load_queue_tasks = []
    widget.image_load_queue_results = []
    widget.tile_image_cache = {}
    widget.canvas_tile_array = []
    return widget


def test_worker_cache_hit_yolunda_photoimage_uretmez(monkeypatch):
    widget = _bare_widget()
    snapshot = widget._server_snapshot()
    assert widget._raw_tile_image_kaydet(
        15,
        1,
        2,
        Image.new("RGB", (256, 256), "red"),
        expected_snapshot=snapshot,
    )

    photo_calls = []

    def forbidden_photoimage(*args, **kwargs):
        photo_calls.append((args, kwargs))
        raise AssertionError("worker PhotoImage üretti")

    monkeypatch.setattr(hassas_zoom.ImageTk, "PhotoImage", forbidden_photoimage)
    result = []
    worker = threading.Thread(
        target=lambda: result.append(
            widget.request_image(15, 1, 2, expected_snapshot=snapshot)
        )
    )
    worker.start()
    worker.join(timeout=2)

    assert not worker.is_alive()
    assert len(result) == 1
    assert isinstance(result[0], Image.Image)
    assert photo_calls == []


def test_eski_render_generation_sonucu_canvas_tile_guncellemez():
    widget = _bare_widget()
    widget._fractional_render_generation = 2

    class FakeTile:
        tile_name_position = (1, 2)
        source_zoom = 15
        image = object()

        def __init__(self):
            self.set_image_calls = 0

        def set_image(self, image):
            self.set_image_calls += 1

    tile = FakeTile()
    stale_token = (widget._server_snapshot(), 1, 15, 1, 2)
    widget.canvas_tile_array = [[tile]]
    widget.image_load_queue_results = [
        (stale_token, tile, Image.new("RGB", (256, 256), "blue")),
    ]
    widget.running = False

    widget.update_canvas_tile_images()

    assert tile.set_image_calls == 0
    assert widget.image_load_queue_results == []


def test_placeholder_raw_cacheteyse_pan_sirasinda_ana_threadde_guncellenir():
    widget = _bare_widget()
    placeholder = object()
    widget.not_loaded_tile_image = placeholder
    raw_image = Image.new("RGB", (256, 256), "orange")
    snapshot = widget._server_snapshot()
    widget._raw_tile_image_kaydet(
        15,
        1,
        2,
        raw_image,
        expected_snapshot=snapshot,
    )

    class FakeTile:
        tile_name_position = (1, 2)
        source_zoom = 15

        def __init__(self):
            self.image = placeholder
            self.applied_images = []

        def set_image(self, image):
            self.applied_images.append(image)
            self.image = image

    tile = FakeTile()
    widget.canvas_tile_array = [[tile]]

    widget._queue_visible_unloaded_tiles(snapshot, widget._fractional_render_generation)

    assert len(tile.applied_images) == 1
    assert tile.image is not placeholder
    assert isinstance(tile.applied_images[0], Image.Image)
    assert widget.image_load_queue_tasks == []


def test_gecici_http_hatasi_empty_karo_cachelemez_ve_sonraki_istek_basarir(monkeypatch):
    widget = _bare_widget()
    snapshot = widget._server_snapshot()
    payload_buffer = io.BytesIO()
    Image.new("RGB", (256, 256), "purple").save(payload_buffer, format="PNG")
    payload = payload_buffer.getvalue()
    responses = []
    request_count = 0

    class TransientFailureResponse:
        raw = io.BytesIO()
        closed = False

        def raise_for_status(self):
            raise requests.exceptions.Timeout("temporary timeout")

        def close(self):
            self.closed = True
            responses.append(self)

    class SuccessResponse:
        def __init__(self):
            self.raw = io.BytesIO(payload)
            self.closed = False

        def raise_for_status(self):
            return None

        def close(self):
            self.closed = True
            responses.append(self)
            self.raw.close()

    def fake_get(_url, **_kwargs):
        nonlocal request_count
        request_count += 1
        if request_count == 1:
            return TransientFailureResponse()
        return SuccessResponse()

    monkeypatch.setattr(hassas_zoom.requests, "get", fake_get)

    assert widget.request_image(15, 1, 2, expected_snapshot=snapshot) is None
    assert len(widget._fractional_raw_tile_cache) == 0

    loaded = widget.request_image(15, 1, 2, expected_snapshot=snapshot)

    assert isinstance(loaded, Image.Image)
    assert len(widget._fractional_raw_tile_cache) == 1
    assert request_count == 2
    assert all(response.closed for response in responses)


def test_raw_ve_scaled_photo_cache_lru_sinirini_korur(monkeypatch):
    widget = _bare_widget()
    snapshot = widget._server_snapshot()
    raw = Image.new("RGB", (1, 1), "white")

    for x in range(RAW_TILE_CACHE_LIMIT + 25):
        assert widget._raw_tile_image_kaydet(
            15,
            x,
            0,
            raw,
            expected_snapshot=snapshot,
        )

    assert len(widget._fractional_raw_tile_cache) == RAW_TILE_CACHE_LIMIT
    assert widget._raw_cache_key(snapshot, 15, 0, 0) not in widget._fractional_raw_tile_cache

    monkeypatch.setattr(
        hassas_zoom.ImageTk,
        "PhotoImage",
        lambda image, **_kwargs: image.size,
    )
    for x in range(SCALED_PHOTO_CACHE_LIMIT + 25):
        widget._raw_tile_image_kaydet(
            15,
            x,
            1,
            raw,
            expected_snapshot=snapshot,
        )
        widget._tile_photo_for(15, x, 1)

    assert len(widget._fractional_photo_cache) == SCALED_PHOTO_CACHE_LIMIT


def test_normal_pan_generation_ve_photo_lruyu_gereksiz_sifirlamaz(monkeypatch):
    widget = _bare_widget()
    placeholder = object()
    widget.not_loaded_tile_image = placeholder

    class LoadedTile:
        source_zoom = 15
        tile_name_position = (1, 2)
        image = object()

    widget.canvas_tile_array = [[LoadedTile()]]
    widget._fractional_photo_cache["kept"] = object()
    initial_generation = widget._fractional_render_generation

    from tkintermapview.map_widget import TkinterMapView

    monkeypatch.setattr(
        TkinterMapView,
        "draw_move",
        lambda self, called_after_zoom=False: None,
    )

    KesirliZoomHaritaView.draw_move(widget)
    KesirliZoomHaritaView.draw_move(widget)

    assert widget._fractional_render_generation == initial_generation
    assert "kept" in widget._fractional_photo_cache

    widget._begin_render_generation()
    assert widget._fractional_render_generation == initial_generation + 1
    assert widget._fractional_photo_cache == {}


def test_server_degisimindeki_eski_http_yaniti_cachee_yazilmaz(monkeypatch):
    widget = _bare_widget()
    old_snapshot = widget._server_snapshot()
    widget._fractional_photo_cache["old-server-photo"] = object()
    payload_buffer = io.BytesIO()
    Image.new("RGB", (256, 256), "green").save(payload_buffer, format="PNG")
    payload = payload_buffer.getvalue()
    request_started = threading.Event()
    release_request = threading.Event()
    responses = []
    request_kwargs = []

    class FakeResponse:
        def __init__(self):
            self.raw = io.BytesIO(payload)
            self.closed = False

        def raise_for_status(self):
            return None

        def close(self):
            self.closed = True
            responses.append(self)
            self.raw.close()

    def fake_get(url, **kwargs):
        request_kwargs.append((url, kwargs))
        request_started.set()
        assert release_request.wait(timeout=2)
        return FakeResponse()

    monkeypatch.setattr(hassas_zoom.requests, "get", fake_get)

    result = []
    worker = threading.Thread(
        target=lambda: result.append(
            widget.request_image(15, 1, 2, expected_snapshot=old_snapshot)
        )
    )
    worker.start()
    assert request_started.wait(timeout=2)

    from tkintermapview.map_widget import TkinterMapView

    def fake_base_set_tile_server(self, tile_server, tile_size=256, max_zoom=19):
        self.tile_server = tile_server
        self.tile_size = tile_size
        self.max_zoom = max_zoom

    monkeypatch.setattr(TkinterMapView, "set_tile_server", fake_base_set_tile_server)
    KesirliZoomHaritaView.set_tile_server(
        widget,
        "new/{z}/{x}/{y}.png",
        tile_size=256,
        max_zoom=22,
    )
    release_request.set()
    worker.join(timeout=2)

    assert not worker.is_alive()
    assert result == [None]
    assert len(widget._fractional_raw_tile_cache) == 0
    assert widget._fractional_photo_cache == {}
    assert responses and all(response.closed for response in responses)
    assert request_kwargs[0][1]["timeout"] == (5.0, 20.0)
