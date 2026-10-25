filter data sku agar sama dengan sebelumnya
AA ganti nama file tanggal hari sebelumnya
=COUNTIF('21-10-2026 STK AA.xls'!Table1[#Data];A5)
NPI ganti nama file tanggal hari sebelumnya
=COUNTIF('21-10-2026 STK NPI.xls'!$A$4:$H$390;A5)

kondisi bold PT produk
=NOT(ISNUMBER(SEARCH("-"; $A5)))
