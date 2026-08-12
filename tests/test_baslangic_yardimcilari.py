from arayuz_yardimcilari import ArayuzYardimcilari
import sys
from types import SimpleNamespace

from bagimlilik_kontrol import surum_yeterli_mi, tkinter_kontrol_et


class SahteUygulama:
    pass


class SahteTree:
    def __init__(self, items, secili):
        self.items = list(items)
        self.secili = list(secili)

    def selection(self):
        return tuple(self.secili)

    def index(self, item):
        return self.items.index(item)

    def get_children(self):
        return tuple(self.items)

    def move(self, item, parent, yeni_index):
        self.items.remove(item)
        self.items.insert(yeni_index, item)

    def item(self, item, **kwargs):
        return None

    def selection_set(self, items):
        self.secili = list(items)

    def focus(self, item):
        return None

    def see(self, item):
        return None


class GorunmeyenHucreTree:
    def __init__(self):
        self.see_cagrildi = False
        self.update_cagrildi = False

    def exists(self, item):
        return True

    def bbox(self, row, column):
        return ""

    def see(self, item):
        self.see_cagrildi = True

    def update_idletasks(self):
        self.update_cagrildi = True


def test_sinirda_secili_satir_blogunun_sirasi_degismez():
    tree = SahteTree(["a", "b", "c", "d"], ["a", "b", "c"])

    ArayuzYardimcilari(SahteUygulama()).tree_secili_satirlari_tasi(tree, -1)

    assert tree.items == ["a", "b", "c", "d"]


def test_secili_satir_blogunun_sirasi_korunarak_tasinir():
    tree = SahteTree(["a", "b", "c", "d"], ["b", "c"])

    ArayuzYardimcilari(SahteUygulama()).tree_secili_satirlari_tasi(tree, 1)

    assert tree.items == ["a", "d", "b", "c"]


def test_paket_surumu_sayisal_olarak_karsilastirilir():
    assert surum_yeterli_mi("1.10.1", "1.9")
    assert surum_yeterli_mi("3.0.0+yerel", "3.0.0")
    assert not surum_yeterli_mi("1.9.9", "1.10.0")
    assert not surum_yeterli_mi("bilinmiyor", "1.0")


def test_gorunmeyen_hucre_bos_bbox_ile_hata_uretmez():
    tree = GorunmeyenHucreTree()

    sonuc = ArayuzYardimcilari(SahteUygulama()).hucre_duzenle(
        None,
        tree,
        set_row="satir-1",
        set_col="#1",
    )

    assert sonuc is None
    assert tree.see_cagrildi
    assert tree.update_cagrildi


def test_tkinter_kontrolu_gercek_tk_penceresi_yontemini_kullanir(monkeypatch):
    cagrilar = []

    class SahteTk:
        def __init__(self):
            self.tk = self

        def withdraw(self):
            cagrilar.append("withdraw")

        def update_idletasks(self):
            cagrilar.append("update")

        def call(self, *args):
            assert args == ("info", "patchlevel")
            return "8.6.12"

        def destroy(self):
            cagrilar.append("destroy")

    monkeypatch.setitem(sys.modules, "tkinter", SimpleNamespace(Tk=SahteTk))

    assert tkinter_kontrol_et() is None
    assert cagrilar == ["withdraw", "update", "destroy"]


def test_tkinter_kontrolu_pencere_hatasini_aciklar(monkeypatch):
    class BozukTk:
        def __init__(self):
            raise RuntimeError("Tk DLL yuklenemedi")

    monkeypatch.setitem(sys.modules, "tkinter", SimpleNamespace(Tk=BozukTk))

    sonuc = tkinter_kontrol_et()

    assert sonuc == "Tk kullanilamiyor: Tk DLL yuklenemedi"
