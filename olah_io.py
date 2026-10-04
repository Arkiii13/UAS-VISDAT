#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
olah_io.py — Pra-pemrosesan Tabel Input-Output Indonesia 2020 (BPS) untuk visualisasi D3.js

Masukan : Tabel I-O Transaksi Total Atas Dasar Harga Dasar (185 Produk), 2020 (.xls)
          [opsional] Tabel I-O Transaksi Domestik Atas Dasar Harga Dasar (17 Produk), 2020 (.xls)
            -> dipakai hanya untuk memvalidasi pemetaan 185 -> 17 kategori.
Keluaran: folder data/ berisi
  nodes.json / nodes.csv      185 node + metrik (strength, betweenness, komunitas, BL, FL, kategori)
  links_semua.json            semua edge i->j (tanpa diagonal), lengkap dengan persentil bobot
  links_ringan.json           subset edge (top-k keluar/masuk per node) untuk tampilan awal graf
  hierarki.json               pohon kelompok -> 17 kategori -> 185 produk (untuk treemap/sunburst)
  flows17.json                aliran 17 x 17 (nodes+links untuk Sankey, matriks untuk chord)
  matriks_185.csv             matriks transaksi antara 185 x 185 (untuk adjacency matrix)
  pemetaan_185_ke_17.csv      pemetaan produk -> kategori -> kelompok
  metadata.json               sumber, satuan, parameter, dan hasil pemeriksaan data

Contoh:
  python olah_io.py \
      --xls185 "indonesia-input-output-table-total-transactions-based-on-basic-prices--185-products-.xls" \
      --xls17  "tabel-input-output-indonesia-transaksi-domestik-atas-dasar-harga-dasar--17-produk-.xls" \
      --out data

Dependensi: pandas, numpy, networkx (>= 2.8), xlrd (>= 2.0.1)
"""
import argparse
import json
import sys
from datetime import date
from pathlib import Path

import networkx as nx
import numpy as np
import pandas as pd

SUMBER = "Sumber: BPS"
JUDUL_185 = ("Tabel Input-Output Indonesia Transaksi Total Atas Dasar Harga Dasar "
             "(185 Produk), 2020")
SATUAN = "Juta rupiah"

# ---------------------------------------------------------------------------
# Pemetaan 185 produk -> 17 kategori lapangan usaha
# Disusun dari urutan kode produk, lalu DIVALIDASI dengan membandingkan jumlah
# "Output Domestik Harga Dasar" per kategori terhadap tabel 17 produk (lihat cek_17).
# Nyatakan di makalah sebagai pengelompokan penulis.
# ---------------------------------------------------------------------------
KATEGORI_17 = {
    1: "Pertanian, Kehutanan, dan Perikanan",
    2: "Pertambangan dan Penggalian",
    3: "Industri Pengolahan",
    4: "Pengadaan Listrik dan Gas",
    5: "Pengadaan Air, Pengelolaan Sampah, Limbah dan Daur Ulang",
    6: "Konstruksi",
    7: "Perdagangan Besar dan Eceran; Reparasi Mobil dan Sepeda Motor",
    8: "Transportasi dan Pergudangan",
    9: "Penyediaan Akomodasi dan Makan Minum",
    10: "Informasi dan Komunikasi",
    11: "Jasa Keuangan dan Asuransi",
    12: "Real Estate",
    13: "Jasa Perusahaan",
    14: "Administrasi Pemerintahan, Pertahanan dan Jaminan Sosial Wajib",
    15: "Jasa Pendidikan",
    16: "Jasa Kesehatan dan Kegiatan Sosial",
    17: "Jasa Lainnya",
}
# (kategori, kode_awal, kode_akhir)
RENTANG = [
    (1, 1, 36), (2, 37, 52), (3, 53, 144), (4, 145, 146), (5, 147, 148),
    (6, 149, 153), (7, 154, 156), (8, 157, 163), (9, 164, 165), (10, 166, 169),
    (11, 170, 173), (12, 174, 174), (13, 175, 176), (14, 177, 177),
    (15, 178, 178), (16, 179, 179), (17, 180, 180), (15, 181, 181),
    (16, 182, 182), (17, 183, 185),
]
# Pengelompokan primer-sekunder-tersier: PENGELOMPOKAN PENULIS (pendekatan sektor Clark).
KELOMPOK = {k: "Primer" for k in (1, 2)}
KELOMPOK.update({k: "Sekunder" for k in (3, 4, 5, 6)})
KELOMPOK.update({k: "Tersier" for k in range(7, 18)})


def bangun_pemetaan():
    peta = {}
    for kat, awal, akhir in RENTANG:
        for kode in range(awal, akhir + 1):
            peta[kode] = kat
    assert sorted(peta) == list(range(1, 186)), "Pemetaan harus menutup kode 1..185"
    return peta


# ---------------------------------------------------------------------------
# Membaca tabel
# ---------------------------------------------------------------------------
def _cari_kolom(df, nama):
    """Indeks kolom yang namanya (baris 4) sama dengan `nama`."""
    hit = [i for i in range(df.shape[1]) if str(df.iat[4, i]).strip() == nama]
    if not hit:
        sys.exit(f"Kolom '{nama}' tidak ditemukan; periksa format file.")
    return hit[0]


def baca_tabel(path, n_produk):
    """Baca file BPS (.xls). Struktur: baris 3 = kode, baris 4 = nama, data mulai baris 5;
    kolom 1 = kode, kolom 2 = nama, kolom produk mulai kolom 3."""
    df = pd.read_excel(path, engine="xlrd", header=None)
    judul = str(df.iat[0, 1]).strip()

    kol_antara = _cari_kolom(df, "Total Permintaan Antara")
    kol_output = _cari_kolom(df, "Output Domestik Harga Dasar")
    n_kol = kol_antara - 3
    if n_kol != n_produk:
        sys.exit(f"Jumlah kolom produk = {n_kol}, diharapkan {n_produk}.")

    baris = df.iloc[5:5 + n_produk]
    kode = [int(float(v)) for v in baris.iloc[:, 1]]
    nama = [str(v).strip() for v in baris.iloc[:, 2]]
    nama_kol = [str(v).strip() for v in df.iloc[4, 3:3 + n_produk]]
    if nama != nama_kol:
        print("PERINGATAN: nama baris dan kolom tidak identik (produk x produk).", file=sys.stderr)

    num = lambda s: pd.to_numeric(s, errors="coerce").fillna(0.0).to_numpy(dtype=float)
    Z = baris.iloc[:, 3:3 + n_produk].apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=float)

    # baris "Total Input" (kode 2100)
    kode_baris = df.iloc[:, 1].astype(str).str.replace(r"\.0$", "", regex=True)
    r_input = kode_baris[kode_baris == "2100"].index
    if len(r_input) != 1:
        sys.exit("Baris 'Total Input' (kode 2100) tidak ditemukan.")
    x = num(df.iloc[r_input[0], 3:3 + n_produk])

    return {
        "judul": judul, "kode": kode, "nama": nama, "Z": Z, "x": x,
        "total_antara": num(baris.iloc[:, kol_antara]),
        "output_domestik": num(baris.iloc[:, kol_output]),
        "raw": df,
    }


def periksa_data(t):
    Z, x = t["Z"], t["x"]
    n = Z.shape[0]
    cek = {
        "ukuran_matriks": f"{Z.shape[0]} x {Z.shape[1]}",
        "jumlah_sel": int(Z.size),
        "sel_positif": int((Z > 0).sum()),
        "kepadatan_persen": round(100 * float((Z > 0).mean()), 2),
        "sel_negatif": int((Z < 0).sum()),
        "diagonal_positif": int((np.diag(Z) > 0).sum()),
        "baris_tanpa_penjualan_antara": [t["nama"][i] for i in np.where(Z.sum(1) == 0)[0]],
        "kolom_tanpa_pembelian_antara": [t["nama"][j] for j in np.where(Z.sum(0) == 0)[0]],
        "selisih_maks_total_antara_vs_jumlah_baris": float(np.abs(t["total_antara"] - Z.sum(1)).max()),
        "total_input_nol": int((x <= 0).sum()),
    }
    assert cek["sel_negatif"] == 0, "Ada nilai negatif pada matriks transaksi"
    assert cek["total_input_nol"] == 0, "Ada total input nol; koefisien A tidak terdefinisi"
    assert cek["selisih_maks_total_antara_vs_jumlah_baris"] < 1, "Total antara tidak cocok dengan jumlah baris"
    return cek


def cek_17(path17, output_185, peta):
    """Bandingkan jumlah Output Domestik per kategori (dari 185) dengan tabel 17 produk."""
    df = pd.read_excel(path17, engine="xlrd", header=None)
    kol = _cari_kolom(df, "Output Domestik Harga Dasar")
    ref = pd.to_numeric(df.iloc[5:22, kol]).to_numpy(dtype=float)
    agg = np.zeros(17)
    for i, kode in enumerate(range(1, 186)):
        agg[peta[kode] - 1] += output_185[i]
    selisih = agg - ref
    hasil = [{"kategori": k, "nama": KATEGORI_17[k], "dari_185": float(agg[k - 1]),
              "tabel_17": float(ref[k - 1]), "selisih": float(selisih[k - 1])}
             for k in range(1, 18)]
    return {"cocok": bool(np.abs(selisih).max() < 1), "per_kategori": hasil}


# ---------------------------------------------------------------------------
# Perhitungan metrik
# ---------------------------------------------------------------------------
def hitung_leontief(Z, x):
    n = Z.shape[0]
    A = Z / x[None, :]                       # a_ij = z_ij / x_j
    L = np.linalg.inv(np.eye(n) - A)         # L = (I - A)^-1
    total = L.sum()
    BL = L.sum(axis=0) * n / total           # keterkaitan ke belakang
    FL = L.sum(axis=1) * n / total           # keterkaitan ke depan
    return A, L, BL, FL


def bangun_graf(Z, kode, nama, sertakan_diagonal=False):
    G = nx.DiGraph()
    for i, (k, nm) in enumerate(zip(kode, nama)):
        G.add_node(k, nama=nm)
    ii, jj = np.nonzero(Z > 0)
    for i, j in zip(ii, jj):
        if i == j and not sertakan_diagonal:
            continue
        G.add_edge(kode[i], kode[j], weight=float(Z[i, j]), dist=1.0 / float(Z[i, j]))
    return G


def hitung_metrik_graf(G, resolusi, skala_log, seed):
    s_out = dict(G.out_degree(weight="weight"))
    s_in = dict(G.in_degree(weight="weight"))
    d_out = dict(G.out_degree())
    d_in = dict(G.in_degree())
    btw = nx.betweenness_centrality(G, weight="dist", normalized=True)

    # Louvain pada graf tak berarah (bobot = z_ij + z_ji)
    U = nx.Graph()
    U.add_nodes_from(G.nodes)
    for u, v, w in G.edges(data="weight"):
        bobot = np.log1p(w) if skala_log else w
        if U.has_edge(u, v):
            U[u][v]["weight"] += bobot
        else:
            U.add_edge(u, v, weight=bobot)
    komunitas = nx.community.louvain_communities(U, weight="weight", resolution=resolusi, seed=seed)
    komunitas = sorted(komunitas, key=lambda c: (-len(c), min(c)))   # id 1 = komunitas terbesar
    id_kom = {n: k + 1 for k, c in enumerate(komunitas) for n in c}
    modularitas = nx.community.modularity(U, komunitas, weight="weight")
    return s_out, s_in, d_out, d_in, btw, id_kom, float(modularitas)


# ---------------------------------------------------------------------------
# Penulisan keluaran
# ---------------------------------------------------------------------------
def _json_default(o):
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating,)):
        return float(o)
    if isinstance(o, (np.bool_,)):
        return bool(o)
    raise TypeError(type(o))


def tulis_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, default=_json_default)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xls185", required=True, help="File .xls Tabel I-O 185 produk (transaksi total)")
    ap.add_argument("--xls17", help="File .xls Tabel I-O 17 produk (untuk validasi pemetaan)")
    ap.add_argument("--out", default="data", help="Folder keluaran (default: data)")
    ap.add_argument("--sertakan-diagonal", action="store_true", help="Sertakan self-loop z_ii pada graf")
    ap.add_argument("--top-k", type=int, default=5, help="Edge keluar & masuk teratas per node untuk links_ringan")
    ap.add_argument("--resolusi", type=float, default=1.0, help="Resolusi Louvain")
    ap.add_argument("--louvain-log", action="store_true", help="Bobot Louvain memakai log1p(nilai)")
    ap.add_argument("--seed", type=int, default=42, help="Seed Louvain (reprodusibilitas)")
    a = ap.parse_args()

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    N = 185

    # 1. Baca dan periksa
    t = baca_tabel(a.xls185, N)
    kode, nama, Z, x = t["kode"], t["nama"], t["Z"], t["x"]
    assert kode == list(range(1, N + 1)), "Kode produk harus 1..185 berurutan"
    cek = periksa_data(t)
    print("Judul tabel :", t["judul"])
    print("Pemeriksaan :", json.dumps(cek, ensure_ascii=False, indent=2))

    peta = bangun_pemetaan()
    cek17 = cek_17(a.xls17, t["output_domestik"], peta) if a.xls17 else None
    if cek17 is not None:
        print("Validasi pemetaan 185 -> 17 terhadap tabel 17 produk:",
              "COCOK" if cek17["cocok"] else "TIDAK COCOK (lihat metadata.json)")

    # 2. Leontief, BL, FL
    A, L, BL, FL = hitung_leontief(Z, x)

    # 3. Graf, sentralitas, komunitas
    G = bangun_graf(Z, kode, nama, a.sertakan_diagonal)
    s_out, s_in, d_out, d_in, btw, id_kom, modul = hitung_metrik_graf(G, a.resolusi, a.louvain_log, a.seed)

    # 4. Tabel node
    nodes = pd.DataFrame({
        "id": kode,
        "nama": nama,
        "kategori_kode": [peta[k] for k in kode],
    })
    nodes["kategori"] = nodes["kategori_kode"].map(KATEGORI_17)
    nodes["kelompok"] = nodes["kategori_kode"].map(KELOMPOK)
    nodes["output_domestik"] = t["output_domestik"]
    nodes["total_input"] = x
    nodes["strength_keluar"] = [s_out[k] for k in kode]
    nodes["strength_masuk"] = [s_in[k] for k in kode]
    nodes["degree_keluar"] = [d_out[k] for k in kode]
    nodes["degree_masuk"] = [d_in[k] for k in kode]
    nodes["betweenness"] = [btw[k] for k in kode]
    nodes["komunitas"] = [id_kom[k] for k in kode]
    nodes["BL"] = BL
    nodes["FL"] = FL
    nodes["sektor_kunci"] = (nodes["BL"] > 1) & (nodes["FL"] > 1)
    # urutan baris/kolom adjacency matrix: per komunitas, lalu strength total menurun
    nodes["strength_total"] = nodes["strength_keluar"] + nodes["strength_masuk"]
    urut = nodes.sort_values(["komunitas", "strength_total"], ascending=[True, False]).index
    nodes["urut_matriks"] = 0
    nodes.loc[urut, "urut_matriks"] = np.arange(N)
    nodes.to_csv(out / "nodes.csv", index=False, encoding="utf-8")
    tulis_json(out / "nodes.json", nodes.to_dict(orient="records"))

    # 5. Edge
    links = pd.DataFrame([(u, v, w) for u, v, w in G.edges(data="weight")],
                         columns=["source", "target", "value"])
    links["persentil"] = (links["value"].rank(pct=True) * 100).round(2)   # untuk slider ambang bobot
    links = links.sort_values("value", ascending=False).reset_index(drop=True)
    tulis_json(out / "links_semua.json", links.to_dict(orient="records"))

    top_out = links.groupby("source", group_keys=False).head(a.top_k)
    top_in = links.groupby("target", group_keys=False).head(a.top_k)
    ringan = pd.concat([top_out, top_in]).drop_duplicates(["source", "target"])
    tulis_json(out / "links_ringan.json", ringan.to_dict(orient="records"))

    # 6. Hierarki: kelompok -> kategori -> produk
    pohon = {"name": "Ekonomi Indonesia", "children": []}
    for kel in ("Primer", "Sekunder", "Tersier"):
        n_kel = {"name": kel, "children": []}
        for kat in sorted(k for k, v in KELOMPOK.items() if v == kel):
            daun = nodes[(nodes.kategori_kode == kat) & (nodes.output_domestik > 0)]
            n_kat = {"name": KATEGORI_17[kat], "kategori_kode": kat, "children": [
                {"name": r.nama, "id": int(r.id), "value": float(r.output_domestik),
                 "BL": float(r.BL), "FL": float(r.FL), "komunitas": int(r.komunitas)}
                for r in daun.itertuples()]}
            n_kel["children"].append(n_kat)
        pohon["children"].append(n_kel)
    tulis_json(out / "hierarki.json", pohon)

    # 7. Aliran 17 x 17
    P = np.zeros((N, 17))
    for i, k in enumerate(kode):
        P[i, peta[k] - 1] = 1.0
    F = P.T @ Z @ P                                    # F[a, b] = aliran dari kategori a ke b
    f_nodes = [{"id": k, "nama": KATEGORI_17[k], "kelompok": KELOMPOK[k],
                "internal": float(F[k - 1, k - 1])} for k in range(1, 18)]
    f_links = [{"source": a_ + 1, "target": b_ + 1, "value": float(F[a_, b_])}
               for a_ in range(17) for b_ in range(17) if a_ != b_ and F[a_, b_] > 0]
    tulis_json(out / "flows17.json", {
        "nodes": f_nodes, "links": f_links,
        "chord": {"nama": [KATEGORI_17[k] for k in range(1, 18)], "matriks": F.tolist()},
        "catatan": "Aliran antarkategori (diagonal dikeluarkan dari links; nilainya ada di nodes.internal). "
                   "Sankey memuat siklus (A->B dan B->A): pakai d3-sankey-circular atau filter ambang."})
    pd.DataFrame(F, index=[KATEGORI_17[k] for k in range(1, 18)],
                 columns=[KATEGORI_17[k] for k in range(1, 18)]).to_csv(out / "flows17_matriks.csv", encoding="utf-8")

    # 8. Matriks 185 x 185 dan pemetaan
    pd.DataFrame(Z, index=kode, columns=kode).rename_axis("kode").to_csv(out / "matriks_185.csv", encoding="utf-8")
    nodes[["id", "nama", "kategori_kode", "kategori", "kelompok"]].to_csv(
        out / "pemetaan_185_ke_17.csv", index=False, encoding="utf-8")

    # 9. Metadata
    tulis_json(out / "metadata.json", {
        "sumber": SUMBER, "judul_tabel": t["judul"], "tahun": 2020, "satuan": SATUAN,
        "tanggal_pengolahan": date.today().isoformat(),
        "catatan_jenis_transaksi": "Transaksi TOTAL (termasuk impor). Indeks BL/FL dari tabel total berbeda "
                                   "dari tabel domestik; nyatakan pilihan ini di Metodologi.",
        "parameter": {"sertakan_diagonal": a.sertakan_diagonal, "top_k": a.top_k,
                      "resolusi_louvain": a.resolusi, "louvain_log1p": a.louvain_log, "seed": a.seed},
        "pemeriksaan_data": cek,
        "graf": {"node": G.number_of_nodes(), "edge": G.number_of_edges(),
                 "jumlah_komunitas": int(nodes["komunitas"].nunique()), "modularitas": modul},
        "validasi_pemetaan_17": cek17,
        "pengelompokan": "Primer/sekunder/tersier adalah pengelompokan penulis.",
    })

    print(f"\nSelesai. Graf: {G.number_of_nodes()} node, {G.number_of_edges()} edge, "
          f"{nodes['komunitas'].nunique()} komunitas (modularitas {modul:.3f}).")
    print(f"Edge pada links_ringan.json: {len(ringan)}. Keluaran di: {out.resolve()}")


if __name__ == "__main__":
    main()
