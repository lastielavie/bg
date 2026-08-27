def buat_excel(df_sumber, raw_bytes=None):
  """Workbook: Sheet 1 (Rincian Faktur Penjualan), Sheet 2 (Pivot), Sheet 3..N (Per Cabang)."""
  buf = io.BytesIO()

  d = df_sumber.copy()
  d['TEKNISI'] = (
      d['TEKNISI']
      .fillna('TIDAK ADA TEKNISI')
      .astype(str)
      .str.strip()
      .replace({
          '': 'TIDAK ADA TEKNISI',
          'nan': 'TIDAK ADA TEKNISI',
          'NaN': 'TIDAK ADA TEKNISI',
          'None': 'TIDAK ADA TEKNISI',
      })
  )
  d['CABANG'] = (
      d['CABANG']
      .fillna('(TANPA CABANG)')
      .astype(str)
      .str.strip()
      .replace({'': '(TANPA CABANG)', 'nan': '(TANPA CABANG)'})
  )

  # 1. Buka template.xlsx bila ada
  if TEMPLATE_PATH.exists():
    wb = openpyxl.load_workbook(TEMPLATE_PATH)

    if raw_bytes and 'Rincian Faktur Penjualan' in wb.sheetnames:
      ws_rincian = wb['Rincian Faktur Penjualan']
      wb_up = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
      ws_up = wb_up.active

      # Hapus data lama (pertahankan Header baris 1 yang berisi penanda 'i')
      if ws_rincian.max_row > 1:
        ws_rincian.delete_rows(2, ws_rincian.max_row)

      # Salin data mentah (baris data tetap kosong pada kolom penyekat)
      for r in range(2, ws_up.max_row + 1):
        row_vals = [ws_up.cell(r, c).value for c in range(1, 47)]
        if any(v is not None for v in row_vals):
          # Kolom 47: Formula Kata Kunci otomatis
          formula_katakunci = f'=_xlfn.IFS(ISNUMBER(SEARCH("Interface", AJ{r})), "Omset Interface", ISNUMBER(SEARCH("Normal", AJ{r})), "Omset Normal", ISNUMBER(SEARCH("Mati Total", AJ{r})), "Omset Mati Total", ISNUMBER(SEARCH("Promo", AJ{r})), "Omset Promo", TRUE, "Omset lainnya")'
          row_vals.append(formula_katakunci)

          ws_rincian.append(row_vals)
  else:
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = 'Rincian Faktur Penjualan'
    wb.create_sheet(title='Pivot')

  # 2. Tambahkan Sheet 3..N per Cabang
  terpakai = set(s.lower() for s in wb.sheetnames)
  for cab in sorted(d['CABANG'].unique()):
    sub = d[d['CABANG'] == cab]
    if sub.empty:
      continue
    dc = rekap_kualifikasi(sub, ['TEKNISI']).copy()
    dc.insert(1, 'Cabang', cab)
    for kol in KOLOM_GAJI:
      dc[kol] = pd.NA

    s_name = _sheet_name(cab, terpakai)
    _tulis_sheet(
        wb, dc, s_name, f'Bagi Hasil Teknisi — Cabang {cab}', kolom_gaji=True
    )

  wb.save(buf)
  buf.seek(0)
  return buf.getvalue()