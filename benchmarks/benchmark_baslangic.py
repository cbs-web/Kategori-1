"""K-1 başlangıç yolunun tekrar üretilebilir soğuk süreç ölçümü."""

from __future__ import annotations

import argparse
import os
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path


KOK = Path(__file__).resolve().parents[1]


def sure_olc(komut, ortam):
    baslangic = time.perf_counter()
    sonuc = subprocess.run(
        komut,
        cwd=KOK,
        env=ortam,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    return time.perf_counter() - baslangic, sonuc.returncode


def ozet(ad, degerler):
    sirali = sorted(degerler)
    print(
        f"{ad}: medyan={statistics.median(sirali) * 1000:.1f} ms; "
        f"min={sirali[0] * 1000:.1f}; p95={sirali[-1] * 1000:.1f}"
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--repeat", type=int, default=7)
    args = parser.parse_args()

    with tempfile.TemporaryDirectory(prefix="k1_benchmark_") as gecici:
        ortam = os.environ.copy()
        for ad in ("LOCALAPPDATA", "APPDATA", "TEMP", "TMP"):
            ortam[ad] = gecici

        # İlk süreç disk/antivirüs önbelleğini ısıtır; rapora dahil edilmez.
        sure_olc([sys.executable, "-c", "import k1"], ortam)
        baslangic = [
            sure_olc([sys.executable, "-c", "import k1"], ortam)[0]
            for _ in range(args.repeat)
        ]
        bagimlilik = [
            sure_olc([sys.executable, "bagimlilik_kontrol.py"], ortam)[0]
            for _ in range(args.repeat)
        ]

    ozet("import k1", baslangic)
    ozet("bağımlılık kontrolü", bagimlilik)
    print("Not: Tk kurulumu bozuksa bağımlılık kontrolü hata kodu döndürebilir; süre yine ölçülür.")


if __name__ == "__main__":
    main()
