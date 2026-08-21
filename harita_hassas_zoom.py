"""Repository-local fractional zoom adapter for tkintermapview 1.29.

tkintermapview keeps ``zoom`` as a float, but its coordinate conversion,
tile selection and vector drawing use ``round(self.zoom)``.  This adapter
keeps that integer as the source tile zoom and applies the fractional part
as a single effective tile scale.  Raster tiles and the package's existing
CanvasPath/CanvasPolygon/CanvasPositionMarker classes therefore share the
same coordinate system without changing site-packages.
"""

from __future__ import annotations

import io
import math
import sqlite3
import threading
import time
from collections import OrderedDict

import requests
from PIL import Image, ImageTk
from tkintermapview.canvas_tile import CanvasTile
from tkintermapview.map_widget import TkinterMapView
from tkintermapview.utility_functions import decimal_to_osm, osm_to_decimal


KESIRLI_ZOOM_ADIMI = 0.25
_ZOOM_KAYIT_HASSASI = 2
RAW_TILE_CACHE_LIMIT = 512
SCALED_PHOTO_CACHE_LIMIT = 128
TILE_CONNECT_TIMEOUT = 5.0
TILE_READ_TIMEOUT = 20.0
TILE_MAX_RETRIES = 3
TILE_RETRY_BASE_DELAY = 0.15


def kesirli_zoom_degerini_duzelt(
    zoom,
    min_zoom=0,
    max_zoom=22,
    adim=KESIRLI_ZOOM_ADIMI,
):
    """Return a finite, clamped zoom snapped to the requested step."""
    try:
        zoom = float(zoom)
        min_zoom = float(min_zoom)
        max_zoom = float(max_zoom)
        adim = float(adim)
    except (TypeError, ValueError):
        zoom = 15.0
        min_zoom, max_zoom, adim = 0.0, 22.0, KESIRLI_ZOOM_ADIMI

    if not math.isfinite(zoom):
        zoom = 15.0
    if not math.isfinite(min_zoom):
        min_zoom = 0.0
    if not math.isfinite(max_zoom):
        max_zoom = 22.0
    if not math.isfinite(adim) or adim <= 0:
        adim = KESIRLI_ZOOM_ADIMI

    if max_zoom < min_zoom:
        max_zoom = min_zoom
    zoom = round(round(zoom / adim) * adim, _ZOOM_KAYIT_HASSASI)
    zoom = max(min_zoom, min(max_zoom, zoom))
    if abs(zoom) < 10 ** (-_ZOOM_KAYIT_HASSASI):
        return 0.0
    return zoom


def kesirli_karo_olcegi(zoom, tile_size=256):
    """Return ``(source_zoom, effective_tile_size)`` for a display zoom."""
    zoom = float(zoom)
    source_zoom = int(round(zoom))
    effective_tile_size = float(tile_size) * (2.0 ** (zoom - source_zoom))
    return source_zoom, effective_tile_size


class _KesirliCanvasTile(CanvasTile):
    """CanvasTile that refreshes a cached source image at current scale."""

    def __init__(self, map_widget, image, tile_name_position, source_zoom):
        super().__init__(map_widget, image, tile_name_position)
        self.source_zoom = source_zoom

    def set_image_and_position(self, image, tile_name_position):
        self.tile_name_position = tile_name_position
        self.source_zoom = self.map_widget._tile_zoom()
        image = self.map_widget._tile_photo_for(
            self.source_zoom,
            tile_name_position[0],
            tile_name_position[1],
            fallback=image,
        )
        self.image = image
        self.draw(image_update=True)

    def set_image(self, image):
        source_zoom = self.source_zoom
        x, y = self.tile_name_position
        image = self.map_widget._tile_photo_for(
            source_zoom,
            x,
            y,
            fallback=image,
        )
        super().set_image(image)


class KesirliZoomHaritaView(TkinterMapView):
    """tkintermapview-compatible map widget with quarter-step zoom."""

    zoom_step = KESIRLI_ZOOM_ADIMI

    def __init__(self, *args, **kwargs):
        self._fractional_main_thread_id = threading.get_ident()
        self._fractional_lock = threading.RLock()
        self._fractional_raw_tile_cache = OrderedDict()
        self._fractional_photo_cache = OrderedDict()
        self._fractional_server_generation = 0
        self._fractional_render_generation = 0
        self._fractional_empty_raw_image = Image.new(
            "RGB",
            (256, 256),
            (190, 190, 190),
        )
        super().__init__(*args, **kwargs)

    def _tile_zoom(self, zoom=None):
        if zoom is None:
            zoom = getattr(self, "zoom", 0)
        try:
            source_zoom = int(round(float(zoom)))
        except (TypeError, ValueError):
            source_zoom = 0
        try:
            max_zoom = int(getattr(self, "max_zoom", 22))
        except (TypeError, ValueError):
            max_zoom = 22
        return max(0, min(max_zoom, source_zoom))

    def _effective_tile_size(self):
        source_zoom, effective_size = kesirli_karo_olcegi(
            getattr(self, "zoom", 0),
            getattr(self, "tile_size", 256),
        )
        source_zoom = max(0, min(self._tile_zoom(), source_zoom))
        if source_zoom != int(round(float(getattr(self, "zoom", 0)))):
            effective_size = float(getattr(self, "tile_size", 256)) * (
                2.0 ** (float(getattr(self, "zoom", 0)) - source_zoom)
            )
        # Canvas/PIL images have integral pixel dimensions.  Use that same
        # display size for the map coordinate span so vector tile boundaries
        # land exactly on raster image boundaries.
        return float(max(1, int(round(effective_size))))

    def _assert_fractional_main_thread(self):
        """Fail loudly if a worker ever reaches a Tk/PhotoImage path."""
        if threading.get_ident() != getattr(
            self,
            "_fractional_main_thread_id",
            threading.get_ident(),
        ):
            raise RuntimeError("PhotoImage işlemleri yalnızca Tk ana threadinde yapılabilir")

    def _server_snapshot_locked(self):
        return (
            int(getattr(self, "_fractional_server_generation", 0)),
            getattr(self, "tile_server", None),
            getattr(self, "overlay_tile_server", None),
            int(getattr(self, "tile_size", 256)),
        )

    def _server_snapshot(self):
        with self._fractional_lock:
            return self._server_snapshot_locked()

    def _server_snapshot_is_current(self, snapshot):
        with self._fractional_lock:
            return snapshot == self._server_snapshot_locked()

    def _raw_cache_key(self, snapshot, source_zoom, x, y):
        # The generation is deliberately part of the key.  A response from
        # the previous base/overlay server can therefore never populate the
        # new server's cache, even if the tile coordinates are identical.
        return (
            int(snapshot[0]),
            int(source_zoom),
            int(x),
            int(y),
        )

    def _tile_photo_cache_key(self, snapshot, source_zoom, x, y, target_size):
        return (
            snapshot,
            int(source_zoom),
            int(x),
            int(y),
            int(target_size),
        )

    def _tile_photo_for(self, source_zoom, x, y, fallback=False):
        """Return a scaled PhotoImage; this method is main-thread-only."""
        self._assert_fractional_main_thread()
        source_zoom, x, y = int(source_zoom), int(x), int(y)
        snapshot = self._server_snapshot()
        raw_key = self._raw_cache_key(snapshot, source_zoom, x, y)

        with self._fractional_lock:
            raw_image = self._fractional_raw_tile_cache.get(raw_key)
            if raw_image is not None:
                self._fractional_raw_tile_cache.move_to_end(raw_key)
            else:
                raw_image = fallback if isinstance(fallback, Image.Image) else None

            if raw_image is None:
                return fallback

            target_size = max(1, int(round(self._effective_tile_size())))
            cache_key = self._tile_photo_cache_key(
                snapshot,
                source_zoom,
                x,
                y,
                target_size,
            )
            cached = self._fractional_photo_cache.get(cache_key)
            if cached is not None:
                self._fractional_photo_cache.move_to_end(cache_key)
                return cached

        if raw_image.size != (target_size, target_size):
            image = raw_image.resize(
                (target_size, target_size),
                Image.Resampling.LANCZOS,
            )
        else:
            image = raw_image

        # This is the only PhotoImage creation path in the adapter.  Worker
        # threads retain PIL images and enqueue them for this main-thread
        # method through update_canvas_tile_images().
        try:
            photo = ImageTk.PhotoImage(image, master=self)
        except TypeError:
            photo = ImageTk.PhotoImage(image)

        with self._fractional_lock:
            self._fractional_photo_cache[cache_key] = photo
            self._fractional_photo_cache.move_to_end(cache_key)
            while len(self._fractional_photo_cache) > SCALED_PHOTO_CACHE_LIMIT:
                self._fractional_photo_cache.popitem(last=False)
        return photo

    def _fractional_tile_cachesini_temizle(self):
        """Invalidate all server-dependent work and release old images."""
        with self._fractional_lock:
            self._fractional_server_generation = (
                int(getattr(self, "_fractional_server_generation", 0)) + 1
            )
            self._fractional_raw_tile_cache.clear()
            # This is called from the Tk thread (server changes are UI
            # operations), so releasing PhotoImage references here is safe.
            self._fractional_photo_cache.clear()
            if hasattr(self, "tile_image_cache"):
                self.tile_image_cache.clear()
            if hasattr(self, "image_load_queue_tasks"):
                self.image_load_queue_tasks.clear()
            if hasattr(self, "image_load_queue_results"):
                self.image_load_queue_results.clear()

    def set_tile_server(self, tile_server: str, tile_size: int = 256, max_zoom: int = 19):
        # Keep the generation bump and the base-widget URL mutation in one
        # critical section.  A worker can finish an old request afterwards,
        # but its snapshot will fail the current-generation check.
        with self._fractional_lock:
            self._fractional_tile_cachesini_temizle()
            return super().set_tile_server(
                tile_server,
                tile_size=tile_size,
                max_zoom=max_zoom,
            )

    def set_overlay_tile_server(self, overlay_server: str):
        with self._fractional_lock:
            super().set_overlay_tile_server(overlay_server)
            self._fractional_tile_cachesini_temizle()
            if hasattr(self, "canvas_tile_array") and hasattr(self, "canvas"):
                self.draw_initial_array()

    def _raw_cache_get(self, snapshot, source_zoom, x, y):
        key = self._raw_cache_key(snapshot, source_zoom, x, y)
        with self._fractional_lock:
            image = self._fractional_raw_tile_cache.get(key)
            if image is not None:
                self._fractional_raw_tile_cache.move_to_end(key)
            return image

    def get_tile_image_from_cache(
        self,
        zoom: int,
        x: int,
        y: int,
        expected_snapshot=None,
    ):
        """Return a PIL tile only; never create a PhotoImage on a worker."""
        snapshot = self._server_snapshot()
        if expected_snapshot is not None and snapshot != expected_snapshot:
            return False
        image = self._raw_cache_get(snapshot, int(zoom), int(x), int(y))
        return image if image is not None else False

    def _raw_tile_image_kaydet(
        self,
        source_zoom,
        x,
        y,
        image,
        expected_snapshot=None,
    ):
        if not isinstance(image, Image.Image):
            return False
        snapshot = expected_snapshot or self._server_snapshot()
        key = self._raw_cache_key(snapshot, source_zoom, x, y)
        with self._fractional_lock:
            if snapshot != self._server_snapshot_locked():
                return False
            self._fractional_raw_tile_cache[key] = image.copy()
            self._fractional_raw_tile_cache.move_to_end(key)
            while len(self._fractional_raw_tile_cache) > RAW_TILE_CACHE_LIMIT:
                self._fractional_raw_tile_cache.popitem(last=False)
        return True

    def _empty_raw_tile(self, tile_size=None):
        tile_size = int(tile_size or getattr(self, "tile_size", 256))
        if tile_size == self._fractional_empty_raw_image.width:
            return self._fractional_empty_raw_image.copy()
        return Image.new("RGB", (tile_size, tile_size), (190, 190, 190))

    @staticmethod
    def _tile_url(template, source_zoom, x, y):
        return (
            template.replace("{x}", str(x))
            .replace("{y}", str(y))
            .replace("{z}", str(source_zoom))
        )

    @staticmethod
    def _download_tile_image(url):
        response = None
        try:
            response = requests.get(
                url,
                stream=True,
                headers={"User-Agent": "TkinterMapView"},
                timeout=(TILE_CONNECT_TIMEOUT, TILE_READ_TIMEOUT),
            )
            response.raise_for_status()
            with Image.open(response.raw) as loaded:
                return loaded.copy()
        finally:
            if response is not None:
                response.close()

    @staticmethod
    def _resize_raw_tile(image, tile_size):
        if image.size == (tile_size, tile_size):
            return image
        return image.resize((tile_size, tile_size), Image.Resampling.LANCZOS)

    def _request_tile_raw(
        self,
        source_zoom,
        x,
        y,
        db_cursor=None,
        expected_snapshot=None,
    ):
        """Load bytes/PIL only and reject results from a changed server."""
        source_zoom, x, y = int(source_zoom), int(x), int(y)
        snapshot = self._server_snapshot()
        if expected_snapshot is not None:
            snapshot = expected_snapshot
            if not self._server_snapshot_is_current(snapshot):
                return None

        cached = self._raw_cache_get(snapshot, source_zoom, x, y)
        if cached is not None:
            return cached

        tile_size = int(snapshot[3])
        image = None
        database_only = bool(getattr(self, "use_database_only", False))

        if db_cursor is not None:
            try:
                db_cursor.execute(
                    "SELECT t.tile_image FROM tiles t "
                    "WHERE t.zoom=? AND t.x=? AND t.y=? AND t.server=?;",
                    (source_zoom, x, y, snapshot[1]),
                )
                result = db_cursor.fetchone()
                if result is not None:
                    with Image.open(io.BytesIO(result[0])) as loaded:
                        image = loaded.copy()
                elif database_only:
                    image = self._empty_raw_tile(tile_size)
            except sqlite3.OperationalError:
                if database_only:
                    image = self._empty_raw_tile(tile_size)
            except Exception:
                if database_only:
                    image = self._empty_raw_tile(tile_size)

        try:
            if image is None:
                if not self._server_snapshot_is_current(snapshot):
                    return None
                image = self._download_tile_image(
                    self._tile_url(snapshot[1], source_zoom, x, y)
                )

            image = self._resize_raw_tile(image, tile_size)

            if snapshot[2] is not None:
                if not self._server_snapshot_is_current(snapshot):
                    return None
                image_overlay = self._download_tile_image(
                    self._tile_url(snapshot[2], source_zoom, x, y)
                )
                image_overlay = self._resize_raw_tile(image_overlay, tile_size).convert("RGBA")
                image = image.convert("RGBA")
                image.paste(image_overlay, (0, 0), image_overlay)

            if not self._raw_tile_image_kaydet(
                source_zoom,
                x,
                y,
                image,
                expected_snapshot=snapshot,
            ):
                return None
            return image
        except (
            Image.UnidentifiedImageError,
            OSError,
            requests.exceptions.RequestException,
        ):
            if not self._server_snapshot_is_current(snapshot):
                return None
            # A timeout, HTTP error (including 429/5xx), or invalid response
            # is not a valid tile.  Do not poison the raw LRU with a gray
            # placeholder; the visible-tile worker will retry with backoff.
            return None
        except Exception:
            if not self._server_snapshot_is_current(snapshot):
                return None
            # Keep unexpected transient transport/decoder failures retryable
            # as well.  A failed request must never become a cache hit.
            return None

    def request_image(
        self,
        zoom: int,
        x: int,
        y: int,
        db_cursor=None,
        expected_snapshot=None,
    ):
        """Compatibility wrapper returning a PIL image, never PhotoImage."""
        return self._request_tile_raw(
            zoom,
            x,
            y,
            db_cursor=db_cursor,
            expected_snapshot=expected_snapshot,
        )

    def _begin_render_generation(self, clear_photos=True):
        """Invalidate queued work and return the token for a new redraw."""
        with self._fractional_lock:
            self._fractional_render_generation = (
                int(getattr(self, "_fractional_render_generation", 0)) + 1
            )
            snapshot = self._server_snapshot_locked()
            if hasattr(self, "image_load_queue_tasks"):
                self.image_load_queue_tasks.clear()
            if hasattr(self, "image_load_queue_results"):
                self.image_load_queue_results.clear()
            # PhotoImage destruction stays on the main thread.  Raw PIL
            # images are intentionally retained in their bounded LRU.  A
            # tile-array change can invalidate tasks without invalidating
            # already-scaled images, so normal pan paths may preserve this
            # LRU by passing clear_photos=False.
            if clear_photos:
                self._fractional_photo_cache.clear()
                if hasattr(self, "tile_image_cache"):
                    self.tile_image_cache.clear()
            return snapshot, self._fractional_render_generation

    @staticmethod
    def _render_task_token(snapshot, render_generation, source_zoom, x, y):
        return (
            snapshot,
            int(render_generation),
            int(source_zoom),
            int(x),
            int(y),
        )

    def _token_parts(self, token):
        """Parse current tokens and tolerate old integer-only test tokens."""
        if not isinstance(token, tuple):
            return None
        if len(token) == 5 and isinstance(token[0], tuple):
            return token
        # Useful compatibility for callers that construct a token with the
        # server generation but not the URL snapshot.
        if len(token) == 5 and isinstance(token[0], int):
            with self._fractional_lock:
                current_snapshot = self._server_snapshot_locked()
            snapshot = (
                int(token[0]),
                current_snapshot[1],
                current_snapshot[2],
                current_snapshot[3],
            )
            return (
                snapshot,
                int(token[1]),
                int(token[2]),
                int(token[3]),
                int(token[4]),
            )
        return None

    def _render_token_is_current(self, token):
        parts = self._token_parts(token)
        if parts is None:
            return False
        snapshot, render_generation, _source_zoom, _x, _y = parts
        with self._fractional_lock:
            return (
                snapshot == self._server_snapshot_locked()
                and int(render_generation) == int(self._fractional_render_generation)
            )

    def _canvas_tile_is_current(self, canvas_tile, source_zoom, x, y):
        for column in getattr(self, "canvas_tile_array", []):
            for tile in column:
                if tile is canvas_tile:
                    return (
                        tuple(getattr(tile, "tile_name_position", ())) == (int(x), int(y))
                        and int(getattr(tile, "source_zoom", source_zoom)) == int(source_zoom)
                    )
        return False

    def _canvas_tile_array_identity(self):
        return tuple(
            tuple(id(canvas_tile) for canvas_tile in column)
            for column in getattr(self, "canvas_tile_array", [])
        )

    def _queue_tile_task(self, canvas_tile, snapshot=None, render_generation=None):
        if snapshot is None:
            snapshot = self._server_snapshot()
        if render_generation is None:
            with self._fractional_lock:
                render_generation = self._fractional_render_generation
        source_zoom = self._tile_zoom()
        x, y = canvas_tile.tile_name_position
        token = self._render_task_token(
            snapshot,
            render_generation,
            source_zoom,
            x,
            y,
        )
        with self._fractional_lock:
            for queued_task in self.image_load_queue_tasks:
                queued_token, queued_tile = queued_task[0], queued_task[1]
                if queued_token == token and queued_tile is canvas_tile:
                    return
            self.image_load_queue_tasks.append((token, canvas_tile))

    def _queue_visible_unloaded_tiles(self, snapshot, render_generation):
        for column in getattr(self, "canvas_tile_array", []):
            for canvas_tile in column:
                if getattr(canvas_tile, "image", None) is not getattr(
                    self,
                    "not_loaded_tile_image",
                    None,
                ):
                    continue
                source_zoom = self._tile_zoom()
                x, y = canvas_tile.tile_name_position
                raw_image = self.get_tile_image_from_cache(
                    source_zoom,
                    x,
                    y,
                    expected_snapshot=snapshot,
                )
                token = self._render_task_token(
                    snapshot,
                    render_generation,
                    source_zoom,
                    x,
                    y,
                )
                if raw_image is False:
                    self._queue_tile_task(canvas_tile, snapshot, render_generation)
                elif (
                    self._render_token_is_current(token)
                    and self._canvas_tile_is_current(
                        canvas_tile,
                        source_zoom,
                        x,
                        y,
                    )
                ):
                    # A worker may have filled the raw cache after the
                    # previous result/task queues were cleared.  Apply that
                    # ready image immediately on Tk's main thread instead of
                    # leaving the placeholder until another full redraw.
                    canvas_tile.set_image(raw_image)

    def pre_cache(self):
        """Pre-cache PIL tiles only; this worker never touches Tk."""
        last_pre_cache_position = None
        last_snapshot = None
        radius = 1
        db_connection = None
        db_cursor = None
        try:
            database_path = getattr(self, "database_path", None)
            if database_path is not None:
                db_connection = sqlite3.connect(database_path)
                db_cursor = db_connection.cursor()

            while getattr(self, "running", False):
                with self._fractional_lock:
                    position = getattr(self, "pre_cache_position", None)
                    snapshot = self._server_snapshot_locked()
                    source_zoom = self._tile_zoom()

                if position != last_pre_cache_position or snapshot != last_snapshot:
                    last_pre_cache_position = position
                    last_snapshot = snapshot
                    radius = 1

                if position is not None and radius <= 8:
                    center_x, center_y = int(round(position[0])), int(round(position[1]))
                    positions = []
                    for x in range(center_x - radius, center_x + radius + 1):
                        positions.append((x, center_y + radius))
                        positions.append((x, center_y - radius))
                    for y in range(center_y - radius, center_y + radius + 1):
                        positions.append((center_x + radius, y))
                        positions.append((center_x - radius, y))

                    for x, y in positions:
                        if self._raw_cache_get(snapshot, source_zoom, x, y) is None:
                            self._request_tile_raw(
                                source_zoom,
                                x,
                                y,
                                db_cursor=db_cursor,
                                expected_snapshot=snapshot,
                            )
                    radius += 1
                else:
                    time.sleep(0.1)
        finally:
            if db_connection is not None:
                db_connection.close()

    def load_images_background(self):
        """Load queued tiles as PIL images and enqueue them for Tk."""
        db_connection = None
        db_cursor = None
        try:
            database_path = getattr(self, "database_path", None)
            if database_path is not None:
                db_connection = sqlite3.connect(database_path)
                db_cursor = db_connection.cursor()

            while getattr(self, "running", False):
                with self._fractional_lock:
                    if self.image_load_queue_tasks:
                        task = self.image_load_queue_tasks.pop()
                    else:
                        task = None

                if task is None:
                    time.sleep(0.01)
                    continue

                token, canvas_tile = task
                parts = self._token_parts(token)
                if parts is None or not self._render_token_is_current(token):
                    continue
                snapshot, _render_generation, source_zoom, x, y = parts
                image = self.get_tile_image_from_cache(
                    source_zoom,
                    x,
                    y,
                    expected_snapshot=snapshot,
                )
                if image is False:
                    image = self.request_image(
                        source_zoom,
                        x,
                        y,
                        db_cursor=db_cursor,
                        expected_snapshot=snapshot,
                    )
                retry_count = 0
                while (
                    image is None
                    and retry_count < TILE_MAX_RETRIES
                    and self._render_token_is_current(token)
                ):
                    # Keep transient failures retryable without a tight
                    # loop.  The task remains local to this worker, so a
                    # single tile has a bounded retry budget and cannot
                    # duplicate itself in the shared queue.
                    time.sleep(TILE_RETRY_BASE_DELAY * (2 ** retry_count))
                    retry_count += 1
                    image = self.request_image(
                        source_zoom,
                        x,
                        y,
                        db_cursor=db_cursor,
                        expected_snapshot=snapshot,
                    )
                if image is None or not isinstance(image, Image.Image):
                    continue
                if not self._render_token_is_current(token):
                    continue

                with self._fractional_lock:
                    if self._render_token_is_current(token):
                        self.image_load_queue_results.append(
                            (token, canvas_tile, image)
                        )
        finally:
            if db_connection is not None:
                db_connection.close()

    def update_canvas_tile_images(self):
        """Apply worker results on Tk's main thread only."""
        self._assert_fractional_main_thread()
        while True:
            with self._fractional_lock:
                if not self.image_load_queue_results:
                    result = None
                else:
                    result = self.image_load_queue_results.pop(0)
            if result is None:
                break

            token, canvas_tile, image = result
            parts = self._token_parts(token)
            if parts is None or not self._render_token_is_current(token):
                continue
            _snapshot, _render_generation, source_zoom, x, y = parts
            if not self._canvas_tile_is_current(canvas_tile, source_zoom, x, y):
                continue
            canvas_tile.set_image(image)

        if getattr(self, "running", False):
            self.after(10, self.update_canvas_tile_images)

    def _new_fractional_tile(self, image, tile_name_position):
        if isinstance(image, Image.Image):
            image = self._tile_photo_for(
                self._tile_zoom(),
                tile_name_position[0],
                tile_name_position[1],
                fallback=image,
            )
        return _KesirliCanvasTile(
            self,
            image,
            tile_name_position,
            self._tile_zoom(),
        )

    def draw_initial_array(self):
        snapshot, render_generation = self._begin_render_generation()
        x_tile_range = max(
            1,
            math.ceil(self.lower_right_tile_pos[0])
            - math.floor(self.upper_left_tile_pos[0]),
        )
        y_tile_range = max(
            1,
            math.ceil(self.lower_right_tile_pos[1])
            - math.floor(self.upper_left_tile_pos[1]),
        )
        upper_left_x = math.floor(self.upper_left_tile_pos[0])
        upper_left_y = math.floor(self.upper_left_tile_pos[1])
        source_zoom = self._tile_zoom()

        for column in self.canvas_tile_array:
            for tile in column:
                tile.delete()
        self.canvas_tile_array = []

        for x_pos in range(x_tile_range):
            column = []
            for y_pos in range(y_tile_range):
                tile_name_position = (upper_left_x + x_pos, upper_left_y + y_pos)
                image = self.get_tile_image_from_cache(
                    source_zoom,
                    *tile_name_position,
                )
                if image is False:
                    image = self.not_loaded_tile_image
                    tile = self._new_fractional_tile(image, tile_name_position)
                    self._queue_tile_task(tile, snapshot, render_generation)
                else:
                    tile = self._new_fractional_tile(image, tile_name_position)
                column.append(tile)
            self.canvas_tile_array.append(column)

        for column in self.canvas_tile_array:
            for tile in column:
                tile.draw()
        for marker in self.canvas_marker_list:
            marker.draw()
        for path in self.canvas_path_list:
            path.draw()
        for polygon in self.canvas_polygon_list:
            polygon.draw()

        self.pre_cache_position = (
            round((self.upper_left_tile_pos[0] + self.lower_right_tile_pos[0]) / 2),
            round((self.upper_left_tile_pos[1] + self.lower_right_tile_pos[1]) / 2),
        )

    def insert_row(self, insert: int, y_name_position: int):
        source_zoom = self._tile_zoom()
        snapshot = self._server_snapshot()
        with self._fractional_lock:
            render_generation = self._fractional_render_generation
        for x_pos in range(len(self.canvas_tile_array)):
            tile_name_position = (
                self.canvas_tile_array[x_pos][0].tile_name_position[0],
                y_name_position,
            )
            image = self.get_tile_image_from_cache(source_zoom, *tile_name_position)
            if image is False:
                image = self.not_loaded_tile_image
                tile = self._new_fractional_tile(image, tile_name_position)
                self._queue_tile_task(tile, snapshot, render_generation)
            else:
                tile = self._new_fractional_tile(image, tile_name_position)
            tile.draw()
            self.canvas_tile_array[x_pos].insert(insert, tile)

    def insert_column(self, insert: int, x_name_position: int):
        if not self.canvas_tile_array:
            return
        source_zoom = self._tile_zoom()
        snapshot = self._server_snapshot()
        with self._fractional_lock:
            render_generation = self._fractional_render_generation
        column = []
        for y_pos in range(len(self.canvas_tile_array[0])):
            tile_name_position = (
                x_name_position,
                self.canvas_tile_array[0][y_pos].tile_name_position[1],
            )
            image = self.get_tile_image_from_cache(source_zoom, *tile_name_position)
            if image is False:
                image = self.not_loaded_tile_image
                tile = self._new_fractional_tile(image, tile_name_position)
                self._queue_tile_task(tile, snapshot, render_generation)
            else:
                tile = self._new_fractional_tile(image, tile_name_position)
            tile.draw()
            column.append(tile)
        self.canvas_tile_array.insert(insert, column)

    def draw_zoom(self):
        self.draw_initial_array()

    def draw_move(self, called_after_zoom: bool = False):
        before_snapshot = self._server_snapshot()
        before_source_zoom = self._tile_zoom()
        before_tile_array = self._canvas_tile_array_identity()
        super().draw_move(called_after_zoom=called_after_zoom)

        after_snapshot = self._server_snapshot()
        after_source_zoom = self._tile_zoom()
        after_tile_array = self._canvas_tile_array_identity()
        source_or_server_changed = (
            before_snapshot != after_snapshot
            or before_source_zoom != after_source_zoom
        )
        tile_array_changed = before_tile_array != after_tile_array

        if source_or_server_changed or tile_array_changed:
            # Existing tile images remain useful when a pan inserts/removes
            # an edge column/row.  Source/server changes, however, require
            # releasing scaled images from the old scale/server.
            snapshot, render_generation = self._begin_render_generation(
                clear_photos=source_or_server_changed,
            )
        else:
            snapshot = after_snapshot
            with self._fractional_lock:
                render_generation = self._fractional_render_generation
        self._queue_visible_unloaded_tiles(snapshot, render_generation)

    def set_zoom(self, zoom: float, relative_pointer_x: float = 0.5, relative_pointer_y: float = 0.5):
        old_zoom = getattr(self, "zoom", 0.0)
        old_source_zoom = self._tile_zoom(old_zoom)
        try:
            pointer_x = float(relative_pointer_x)
            pointer_y = float(relative_pointer_y)
        except (TypeError, ValueError):
            pointer_x, pointer_y = 0.5, 0.5
        pointer_x = max(0.0, min(1.0, pointer_x))
        pointer_y = max(0.0, min(1.0, pointer_y))

        try:
            tile_width = self.lower_right_tile_pos[0] - self.upper_left_tile_pos[0]
            tile_height = self.lower_right_tile_pos[1] - self.upper_left_tile_pos[1]
            current_tile_x = self.upper_left_tile_pos[0] + tile_width * pointer_x
            current_tile_y = self.upper_left_tile_pos[1] + tile_height * pointer_y
            current_deg_position = osm_to_decimal(
                current_tile_x,
                current_tile_y,
                old_source_zoom,
            )
        except (AttributeError, IndexError, TypeError, ValueError, ZeroDivisionError):
            current_deg_position = (0.0, 0.0)

        min_zoom = getattr(self, "min_zoom", 0)
        max_zoom = getattr(self, "max_zoom", 22)
        self.zoom = kesirli_zoom_degerini_duzelt(
            zoom,
            min_zoom=min_zoom,
            max_zoom=max_zoom,
            adim=self.zoom_step,
        )
        # Source PIL tiles remain reusable, while scaled PhotoImage objects
        # from the previous quarter-step are no longer needed after redraw.
        getattr(self, "_fractional_photo_cache", {}).clear()
        source_zoom = self._tile_zoom()
        current_tile_position = decimal_to_osm(
            current_deg_position[0],
            current_deg_position[1],
            source_zoom,
        )
        effective_tile_size = self._effective_tile_size()
        width = max(1.0, float(getattr(self, "width", 1)))
        height = max(1.0, float(getattr(self, "height", 1)))
        self.upper_left_tile_pos = (
            current_tile_position[0] - pointer_x * (width / effective_tile_size),
            current_tile_position[1] - pointer_y * (height / effective_tile_size),
        )
        self.lower_right_tile_pos = (
            current_tile_position[0] + (1 - pointer_x) * (width / effective_tile_size),
            current_tile_position[1] + (1 - pointer_y) * (height / effective_tile_size),
        )
        self.check_map_border_crossing()
        self.last_zoom = self.zoom
        if hasattr(self, "canvas_tile_array"):
            self.draw_initial_array()

    def set_position(self, deg_x, deg_y, text=None, marker=False, **kwargs):
        current_tile_position = decimal_to_osm(
            deg_x,
            deg_y,
            self._tile_zoom(),
        )
        effective_tile_size = self._effective_tile_size()
        width = max(1.0, float(getattr(self, "width", 1)))
        height = max(1.0, float(getattr(self, "height", 1)))
        self.upper_left_tile_pos = (
            current_tile_position[0] - ((width / 2) / effective_tile_size),
            current_tile_position[1] - ((height / 2) / effective_tile_size),
        )
        self.lower_right_tile_pos = (
            current_tile_position[0] + ((width / 2) / effective_tile_size),
            current_tile_position[1] + ((height / 2) / effective_tile_size),
        )
        if marker is True:
            marker_object = self.set_marker(deg_x, deg_y, text, **kwargs)
        else:
            marker_object = None
        self.check_map_border_crossing()
        self.draw_initial_array()
        return marker_object

    def get_position(self):
        source_zoom = self._tile_zoom()
        return osm_to_decimal(
            (self.lower_right_tile_pos[0] + self.upper_left_tile_pos[0]) / 2,
            (self.lower_right_tile_pos[1] + self.upper_left_tile_pos[1]) / 2,
            source_zoom,
        )

    def convert_canvas_coords_to_decimal_coords(self, canvas_x: int, canvas_y: int) -> tuple:
        relative_mouse_x = canvas_x / self.canvas.winfo_width()
        relative_mouse_y = canvas_y / self.canvas.winfo_height()
        tile_mouse_x = self.upper_left_tile_pos[0] + (
            self.lower_right_tile_pos[0] - self.upper_left_tile_pos[0]
        ) * relative_mouse_x
        tile_mouse_y = self.upper_left_tile_pos[1] + (
            self.lower_right_tile_pos[1] - self.upper_left_tile_pos[1]
        ) * relative_mouse_y
        return osm_to_decimal(tile_mouse_x, tile_mouse_y, self._tile_zoom())

    def check_map_border_crossing(self):
        diff_x, diff_y = 0.0, 0.0
        if self.upper_left_tile_pos[0] < 0:
            diff_x += -self.upper_left_tile_pos[0]
        if self.upper_left_tile_pos[1] < 0:
            diff_y += -self.upper_left_tile_pos[1]
        world_size = 2 ** self._tile_zoom()
        if self.lower_right_tile_pos[0] > world_size:
            diff_x -= self.lower_right_tile_pos[0] - world_size
        if self.lower_right_tile_pos[1] > world_size:
            diff_y -= self.lower_right_tile_pos[1] - world_size
        self.upper_left_tile_pos = (
            self.upper_left_tile_pos[0] + diff_x,
            self.upper_left_tile_pos[1] + diff_y,
        )
        self.lower_right_tile_pos = (
            self.lower_right_tile_pos[0] + diff_x,
            self.lower_right_tile_pos[1] + diff_y,
        )

    def mouse_zoom(self, event):
        relative_mouse_x = event.x / max(1, self.width)
        relative_mouse_y = event.y / max(1, self.height)
        if getattr(event, "num", None) == 4:
            direction = 1
        elif getattr(event, "num", None) == 5:
            direction = -1
        elif getattr(event, "delta", 0) > 0:
            direction = 1
        elif getattr(event, "delta", 0) < 0:
            direction = -1
        else:
            return
        self.set_zoom(
            self.zoom + direction * self.zoom_step,
            relative_pointer_x=relative_mouse_x,
            relative_pointer_y=relative_mouse_y,
        )

    def button_zoom_in(self):
        self.set_zoom(self.zoom + self.zoom_step, relative_pointer_x=0.5, relative_pointer_y=0.5)

    def button_zoom_out(self):
        self.set_zoom(self.zoom - self.zoom_step, relative_pointer_x=0.5, relative_pointer_y=0.5)
