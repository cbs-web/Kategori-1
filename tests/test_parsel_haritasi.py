import math
import threading
from collections import OrderedDict
from types import SimpleNamespace

from PIL import Image

from cizimler import parsel_gorunum_hesapla, parsel_noktalari_hashi
from harita_hassas_zoom import (
    KesirliZoomHaritaView,
    kesirli_karo_olcegi,
    kesirli_zoom_degerini_duzelt,
)
from harita_islemleri import (
    HaritaIslemleri,
    _parsel_kadraj_mercator,
    parsel_etkilesimli_kadraj_hesapla,
)
from on_deger_islemleri import OnDegerIslemleri


def test_parsel_gorunumu_poligon_merkezini_bulup_makul_zoom_secer():
    points = [
        (39.52390, 26.11980),
        (39.52390, 26.12030),
        (39.52425, 26.12030),
        (39.52425, 26.11980),
        (39.52390, 26.11980),
    ]

    center, zoom = parsel_gorunum_hesapla(points, 1000, 650)

    assert 39.52390 <= center[0] <= 39.52425
    assert 26.11980 <= center[1] <= 26.12030
    assert 17 <= zoom <= 21
    assert parsel_noktalari_hashi(points) == parsel_noktalari_hashi(list(points))


def test_etkilesimli_kadraj_uzun_parseli_guvenli_tam_zoomla_sigar():
    points = [
        (39.50000, 26.12000),
        (39.50000, 26.12050),
        (39.52500, 26.12050),
        (39.52500, 26.12000),
        (39.50000, 26.12000),
    ]

    width, height, edge = 850, 650, 0.10
    center, zoom, ideal_zoom = parsel_etkilesimli_kadraj_hesapla(points, width, height, edge)

    assert 39.50000 <= center[0] <= 39.52500
    assert 26.12000 <= center[1] <= 26.12050
    assert zoom < ideal_zoom < zoom + 1

    mercator = [_parsel_kadraj_mercator(*point) for point in points]
    min_x, max_x = min(point[0] for point in mercator), max(point[0] for point in mercator)
    min_y, max_y = min(point[1] for point in mercator), max(point[1] for point in mercator)
    center_x, center_y = _parsel_kadraj_mercator(*center)
    assert math.isclose(center_x, (min_x + max_x) / 2.0, abs_tol=1e-12)
    assert math.isclose(center_y, (min_y + max_y) / 2.0, abs_tol=1e-12)

    world_pixels = 256.0 * (2 ** zoom)
    bbox_width = (max_x - min_x) * world_pixels
    bbox_height = (max_y - min_y) * world_pixels
    horizontal_padding = (width - bbox_width) / (2.0 * width)
    vertical_padding = (height - bbox_height) / (2.0 * height)
    assert horizontal_padding >= edge
    assert edge <= vertical_padding <= 0.31
    assert (height - bbox_height * 2.0) / (2.0 * height) < edge


class _SahteHarita:
    min_zoom = 1
    max_zoom = 22

    def __init__(self, width=850, height=650):
        self.width = width
        self.height = height
        self.canvas = self
        self.calls = []
        self.bindings = {}
        self.position = (39.524, 26.120)
        self.zoom = 15

    def update_idletasks(self):
        return None

    def winfo_width(self):
        return self.width

    def winfo_height(self):
        return self.height

    def bind(self, event, callback, add=None):
        self.bindings[event] = callback
        return "binding-1"

    def unbind(self, event, binding_id):
        self.bindings.pop(event, None)

    def set_zoom(self, zoom):
        self.zoom = zoom
        self.calls.append(("zoom", zoom))

    def set_position(self, lat, lon):
        self.position = (lat, lon)
        self.calls.append(("position", lat, lon))

    def get_position(self):
        return self.position


class _SahteDugme:
    def __init__(self):
        self.text = "Parseli Odakla"

    def configure(self, **kwargs):
        self.text = kwargs["text"]


class _SahteRoot:
    def __init__(self):
        self.after_idle_callback = None

    def after_idle(self, callback):
        self.after_idle_callback = callback
        return "idle-1"

    def after_cancel(self, _after_id):
        return None


def test_parseli_odakla_widget_boyutu_hazir_degilse_bekler():
    points = [
        (39.52390, 26.11980),
        (39.52390, 26.12030),
        (39.52425, 26.12030),
        (39.52425, 26.11980),
    ]
    harita = _SahteHarita(width=1, height=1)
    root = _SahteRoot()
    app = SimpleNamespace(root=root, map_widget=harita, yuklu_kml_points=points)

    assert HaritaIslemleri(app).parseli_odakla() is False
    assert not harita.calls

    harita.width, harita.height = 850, 650
    root.after_idle_callback()

    assert harita.calls and harita.calls[0][0] == "zoom"
    assert "<Configure>" not in harita.bindings


def test_parseli_odakla_ikinci_kullanimda_onceki_gorunume_doner():
    points = [
        (39.52390, 26.11980),
        (39.52390, 26.12030),
        (39.52425, 26.12030),
        (39.52425, 26.11980),
    ]
    harita = _SahteHarita()
    dugme = _SahteDugme()
    app = SimpleNamespace(
        map_widget=harita,
        yuklu_kml_points=points,
        btn_parseli_odakla=dugme,
    )
    islem = HaritaIslemleri(app)

    assert islem.parseli_odakla() is True
    assert dugme.text == "Yakın Görünüme Dön"
    assert islem.parseli_odakla() is True
    assert dugme.text == "Parseli Odakla"
    assert harita.position == (39.524, 26.120)
    assert harita.zoom == 15


def test_otomatik_odak_yakin_gorunumu_parsel_merkezinde_bir_ust_zoom_yapar():
    points = [
        (39.50000, 26.12000),
        (39.50000, 26.12050),
        (39.52500, 26.12050),
        (39.52500, 26.12000),
    ]
    harita = _SahteHarita()
    dugme = _SahteDugme()
    app = SimpleNamespace(
        map_widget=harita,
        yuklu_kml_points=points,
        btn_parseli_odakla=dugme,
    )
    islem = HaritaIslemleri(app)

    assert islem.parseli_odakla(otomatik=True) is True
    tam_zoom = harita.zoom
    tam_merkez = harita.position
    assert islem.parseli_odakla() is True
    assert harita.zoom == tam_zoom + 0.25
    assert harita.position == tam_merkez


def test_salt_okunur_acik_widget_listesi_ekleyerek_korunur():
    ilk = object()
    ikinci = object()
    app = SimpleNamespace(salt_okunurda_acik_widgetlar=[ilk])

    sonuc = OnDegerIslemleri(app).salt_okunurda_acik_widget_ekle(ikinci, ilk)

    assert sonuc == [ilk, ikinci]
    assert app.salt_okunurda_acik_widgetlar == [ilk, ikinci]


def test_kesirli_zoom_karo_vektor_olcegini_ayni_tutar():
    assert kesirli_zoom_degerini_duzelt(15.12) == 15.0
    assert kesirli_zoom_degerini_duzelt(15.13) == 15.25

    source_zoom, effective_tile_size = kesirli_karo_olcegi(15.25)
    assert source_zoom == 15
    assert math.isclose(effective_tile_size, 256 * (2 ** 0.25))

    # The adapter stores vector coordinates at source_zoom and maps one
    # source tile to effective_tile_size screen pixels.
    point_a = (39.524, 26.120)
    point_b = (39.524, 26.1201)
    source_a = _parsel_kadraj_mercator(*point_a)
    source_b = _parsel_kadraj_mercator(*point_b)
    source_span = (source_b[0] - source_a[0]) * (2 ** source_zoom)
    screen_width = 800
    source_view_width = screen_width / effective_tile_size
    vector_delta = source_span / source_view_width * screen_width
    raster_delta = source_span * effective_tile_size
    assert math.isclose(vector_delta, raster_delta, rel_tol=1e-12)


def test_kesirli_adapter_raster_karo_goruntusunu_efektif_boyuta_olcekler(monkeypatch):
    widget = object.__new__(KesirliZoomHaritaView)
    widget.zoom = 15.25
    widget.max_zoom = 22
    widget.tile_size = 256
    widget._fractional_raw_tile_cache = {
        (0, 15, 1, 2): Image.new("RGB", (256, 256), "red"),
    }
    widget._fractional_raw_tile_cache = OrderedDict(widget._fractional_raw_tile_cache)
    widget._fractional_photo_cache = OrderedDict()
    widget._fractional_main_thread_id = threading.get_ident()
    widget._fractional_lock = threading.RLock()
    widget._fractional_server_generation = 0
    widget.tile_server = "https://tiles/{z}/{x}/{y}.png"
    widget.overlay_tile_server = None
    monkeypatch.setattr(
        "harita_hassas_zoom.ImageTk.PhotoImage",
        lambda image, **_kwargs: image.size,
    )

    assert widget._tile_photo_for(15, 1, 2) == (304, 304)


def test_kesirli_zoom_widget_ayarinda_ve_kontrollerde_ceyrek_adim_kullanir():
    widget = object.__new__(KesirliZoomHaritaView)
    widget.zoom = 15.0
    widget.min_zoom = 1
    widget.max_zoom = 22
    widget.width = 800
    widget.height = 600
    widget.tile_size = 256
    widget.upper_left_tile_pos = (0.0, 0.0)
    widget.lower_right_tile_pos = (1.0, 1.0)
    widget.last_zoom = 15.0
    widget.draw_initial_array = lambda: None
    widget.check_map_border_crossing = lambda: None

    KesirliZoomHaritaView.set_zoom(widget, 15.13)
    assert widget.zoom == 15.25
    assert math.isclose(
        widget.lower_right_tile_pos[0] - widget.upper_left_tile_pos[0],
        widget.width / round(256 * (2 ** 0.25)),
    )

    calls = []

    def fake_set_zoom(zoom, **kwargs):
        calls.append(zoom)
        widget.zoom = zoom

    widget.set_zoom = fake_set_zoom
    KesirliZoomHaritaView.mouse_zoom(widget, SimpleNamespace(x=400, y=300, delta=120))
    KesirliZoomHaritaView.button_zoom_out(widget)
    assert calls == [15.5, 15.25]
