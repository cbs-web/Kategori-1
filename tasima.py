import math


PHI_SIFIR_TOLERANSI_DERECE = 1e-8
BOWLES_ZEMIN_KATSAYISI = 40.0


class TBDY2018TasimaGucu:
    """TBDY 2018 yüzeysel temel taşıma gücü hesaplama motoru.

    ``b`` ve ``l`` etkili temel boyutlarıdır ve ``b <= l`` olmalıdır. Eğik
    yük, eğimli zemin ve eğimli temel tabanı düzeltmeleri, açıkça farklı
    katsayı verilmediği sürece 1.0 (düşey/merkezî yük ve yatay yüzeyler)
    kabul edilir.
    """

    def __init__(self, c, phi, gamma_n, gamma_sat=None, su_seviyesi=999.0):
        try:
            c = float(c)
            phi = float(phi)
            gamma_n = float(gamma_n)
            gamma_sat = float(gamma_sat) if gamma_sat is not None else gamma_n
            su_seviyesi = float(su_seviyesi)
        except (TypeError, ValueError) as exc:
            raise ValueError("Taşıma gücü girdileri sayısal olmalıdır.") from exc

        girdiler = {
            "kohezyon": c,
            "sürtünme açısı": phi,
            "doğal birim hacim ağırlık": gamma_n,
            "doygun birim hacim ağırlık": gamma_sat,
            "yeraltı su seviyesi": su_seviyesi,
        }
        for ad, deger in girdiler.items():
            if not math.isfinite(float(deger)):
                raise ValueError(f"{ad} sonlu bir sayı olmalıdır.")
        if c < 0:
            raise ValueError("Kohezyon negatif olamaz.")
        if not 0 <= phi <= 45:
            raise ValueError("İçsel sürtünme açısı 0 ile 45 derece arasında olmalıdır.")
        if gamma_n <= 0 or gamma_sat <= 0:
            raise ValueError("Birim hacim ağırlıklar sıfırdan büyük olmalıdır.")
        if gamma_sat < gamma_n:
            raise ValueError("Doygun birim hacim ağırlık doğal birim hacim ağırlıktan küçük olamaz.")
        if gamma_sat <= 9.81:
            raise ValueError("Doygun birim hacim ağırlık suyun birim hacim ağırlığından büyük olmalıdır.")
        if su_seviyesi < 0:
            raise ValueError("Yeraltı su seviyesi negatif olamaz.")

        self.c = c
        self.phi = phi
        self.phi_rad = math.radians(phi)
        self.gamma_n = gamma_n
        self.gamma_sat = gamma_sat
        self.su_seviyesi = su_seviyesi

    def _hesapla_tasima_gucu_faktorleri(self):
        if abs(self.phi) <= PHI_SIFIR_TOLERANSI_DERECE:
            return 5.14, 1.0, 0.0
        tan_phi = math.tan(self.phi_rad)
        sin_phi = math.sin(self.phi_rad)
        # tan²(45° + φ/2) = (1 + sinφ) / (1 - sinφ). Logaritmik
        # biçim ve expm1, φ sıfıra yaklaşırken Nq-1 farkındaki hassasiyet
        # kaybını önler.
        log_nq = math.pi * tan_phi + math.log1p(sin_phi) - math.log1p(-sin_phi)
        nq_eksi_bir = math.expm1(log_nq)
        nq = 1.0 + nq_eksi_bir
        nc = nq_eksi_bir / tan_phi
        ngamma = 2 * nq_eksi_bir * tan_phi
        return nc, nq, ngamma

    def _hesapla_sekil_faktorleri(self, b, l, nc, nq):
        if l == 0:
            l = b
        if abs(self.phi) <= PHI_SIFIR_TOLERANSI_DERECE:
            return 1 + 0.2 * (b / l), 1.0, 1.0
        sc = 1 + (nq / nc) * (b / l)
        sq = 1 + (b / l) * math.tan(self.phi_rad)
        sgamma = 1 - 0.4 * (b / l)
        return sc, sq, sgamma

    def _hesapla_derinlik_faktorleri(self, b, df, nc, nq):
        if b == 0:
            return 1, 1, 1
        k = df / b if df / b <= 1 else math.atan(df / b)
        if abs(self.phi) <= PHI_SIFIR_TOLERANSI_DERECE:
            return 1 + 0.4 * k, 1.0, 1.0
        term = (1 - math.sin(self.phi_rad)) ** 2
        dc = 1 + 0.4 * k
        dq = 1 + 2 * math.tan(self.phi_rad) * term * k
        dgamma = 1.0
        return dc, dq, dgamma

    def _su_etkisi_duzeltmesi(self, df, b):
        gamma_su = 9.81
        if self.su_seviyesi <= df:
            gamma_2 = self.gamma_sat - gamma_su
        elif self.su_seviyesi >= df + b:
            gamma_2 = self.gamma_n
        else:
            d = self.su_seviyesi - df
            gamma_2 = (d * self.gamma_n + (b - d) * (self.gamma_sat - gamma_su)) / b

        if self.su_seviyesi <= 0:
            q = df * (self.gamma_sat - gamma_su)
        elif self.su_seviyesi >= df:
            q = df * self.gamma_n
        else:
            zw = self.su_seviyesi
            q = (zw * self.gamma_n) + ((df - zw) * (self.gamma_sat - gamma_su))

        return gamma_2, q

    def analiz_yap(self, b, l, df, gamma_Rv=1.4, duzeltme_katsayilari=None):
        try:
            b = float(b)
            l = float(l)
            df = float(df)
            gamma_Rv = float(gamma_Rv)
        except (TypeError, ValueError) as exc:
            raise ValueError("Temel geometrisi ve dayanım katsayısı sayısal olmalıdır.") from exc

        for ad, deger in {
            "etkili temel genişliği": b,
            "etkili temel uzunluğu": l,
            "temel derinliği": df,
            "dayanım katsayısı": gamma_Rv,
        }.items():
            if not math.isfinite(float(deger)):
                raise ValueError(f"{ad} sonlu bir sayı olmalıdır.")
        if b <= 0 or l <= 0:
            raise ValueError("Etkili temel boyutları sıfırdan büyük olmalıdır.")
        if b > l:
            raise ValueError("B etkili kısa boyuttur ve L'den büyük olamaz (B <= L).")
        if df < 0:
            raise ValueError("Temel derinliği negatif olamaz.")
        if gamma_Rv <= 0:
            raise ValueError("Dayanım katsayısı sıfırdan büyük olmalıdır.")

        nc, nq, ngamma = self._hesapla_tasima_gucu_faktorleri()
        sc, sq, sgamma = self._hesapla_sekil_faktorleri(b, l, nc, nq)
        dc, dq, dgamma = self._hesapla_derinlik_faktorleri(b, df, nc, nq)
        gamma_eff, q_sur = self._su_etkisi_duzeltmesi(df, b)

        self.Nc, self.Nq, self.Ngamma = nc, nq, ngamma
        self.sc, self.sq, self.sgamma = sc, sq, sgamma
        self.dc, self.dq, self.dgamma = dc, dq, dgamma
        katsayilar = {
            "ic": 1.0, "iq": 1.0, "igamma": 1.0,
            "gc": 1.0, "gq": 1.0, "ggamma": 1.0,
            "bc": 1.0, "bq": 1.0, "bgamma": 1.0,
        }
        if duzeltme_katsayilari:
            bilinmeyen = set(duzeltme_katsayilari) - set(katsayilar)
            if bilinmeyen:
                raise ValueError("Bilinmeyen düzeltme katsayıları: " + ", ".join(sorted(bilinmeyen)))
            for anahtar, deger in duzeltme_katsayilari.items():
                deger = float(deger)
                if not math.isfinite(deger) or not 0 <= deger <= 1:
                    raise ValueError(f"{anahtar} 0 ile 1 arasında sonlu bir sayı olmalıdır.")
                katsayilar[anahtar] = deger

        self.ic, self.iq, self.igamma = katsayilar["ic"], katsayilar["iq"], katsayilar["igamma"]
        self.gc, self.gq, self.ggamma = katsayilar["gc"], katsayilar["gq"], katsayilar["ggamma"]
        self.bc, self.bq, self.bgamma = katsayilar["bc"], katsayilar["bq"], katsayilar["bgamma"]

        qk = (
            self.c * nc * sc * dc * self.ic * self.gc * self.bc
            + q_sur * nq * sq * dq * self.iq * self.gq * self.bq
            + 0.5 * gamma_eff * b * ngamma * sgamma * dgamma * self.igamma * self.ggamma * self.bgamma
        )
        qt = qk / gamma_Rv

        return qk, qt
