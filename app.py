import pandas as pd
from openpyxl import load_workbook


def process_faktur(input_file, template_file, output_file):
    # 1. Baca Data Input & Template
    df_raw = pd.read_excel(input_file)
    wb = load_workbook(template_file)

    # 2. Pembersihan & Penguncian Logika Omset ke 'Nama Teknisi Final'
    # Pastikan nama kolom sesuai dengan file (contoh: 'Teknisi Final' atau 'Nama Teknisi Final')
    col_teknisi = "Nama Teknisi Final"
    col_omset = "Total"  # Ubah sesuai header nilai omset/penjualan di file Anda

    # Bersihkan whitespace
    df_raw[col_teknisi] = df_raw[col_teknisi].astype(str).str.strip()

    # Filter data yang memiliki nama teknisi valid (abaikan NaN atau kosong)
    df_valid = df_raw[
        df_raw[col_teknisi].notna() & (df_raw[col_teknisi] != "")
    ].copy()

    # Rekap Omset dikunci berdasarkan Nama Teknisi Final
    rekap_omset = (
        df_valid.groupby(col_teknisi)[col_omset]
        .agg(Total_Omset="sum", Jumlah_Transaksi="count")
        .reset_index()
    )

    # 3. Penulisan ke Sheet Template (Sheet 1 & Sheet 2)
    # Sheet 1: Rekap Omset per Teknisi Final
    ws1 = wb.worksheets[0]
    # Mulai tulis dari baris data (misal baris 2 setelah header)
    for row_idx, row in rekap_omset.iterrows():
        ws1.cell(row=row_idx + 2, column=1, value=row[col_teknisi])
        ws1.cell(row=row_idx + 2, column=2, value=row["Jumlah_Transaksi"])
        ws1.cell(row=row_idx + 2, column=3, value=row["Total_Omset"])

    # Sheet 2: Detail Transaksi berdasarkan Teknisi Final
    ws2 = wb.worksheets[1]
    for row_idx, row in df_valid.iterrows():
        # Masukkan kolom detail sesuai template Anda
        ws2.append(row.tolist())

    # 4. Tambahkan Sheet 3 (Sheet N / Data Mentah Asli)
    if "Sheet N" in wb.sheetnames:
        ws3 = wb["Sheet N"]
    else:
        ws3 = wb.create_sheet(title="Sheet N")

    # Tulis header dan seluruh baris raw data
    ws3.append(list(df_raw.columns))
    for row in df_raw.itertuples(index=False):
        ws3.append(list(row))

    # 5. Simpan Hasil
    wb.save(output_file)


# Jalankan fungsi
process_faktur(
    input_file="rincian_faktur_penjualan_009mflashsawang_260827141501.xlsx",
    template_file="template.xlsx",
    output_file="hasil_proses_faktur.xlsx",
)