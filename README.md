# 📊 SKU Stock Data Filter

Dokumentasi proses pencocokan data SKU dengan data stok hari sebelumnya menggunakan Microsoft Excel.

## 📌 Tujuan

Memastikan data SKU pada file saat ini sesuai dengan data stok hari sebelumnya untuk kategori **AA** dan **NPI**, serta membedakan baris PT produk menggunakan format **bold**.

## 📂 File Referensi

| Kategori | Nama File | Sumber Data |
|---|---|---|
| AA | `DD-MM-YYYY STK AA.xls` | `Table1[#Data]` |
| NPI | `DD-MM-YYYY STK NPI.xls` | Range `A4:H390` |

`DD-MM-YYYY` mengikuti tanggal data stok hari sebelumnya.

## 🔍 1. Filter SKU — AA

Gunakan rumus `COUNTIF` untuk memeriksa apakah SKU pada sel `A5` terdapat dalam data referensi AA.

```excel
=COUNTIF('21-10-2026 STK AA.xls'!Table1[#Data];A5)
```

**Keterangan:**
- `Table1[#Data]` merupakan sumber data SKU pada file AA.
- `A5` merupakan SKU yang akan dicocokkan.
- Hasil `0` berarti tidak ditemukan kecocokan.
- Hasil lebih dari `0` berarti SKU ditemukan setidaknya satu kali.

## 🔍 2. Filter SKU — NPI

Gunakan rumus berikut untuk mencocokkan SKU dengan data referensi NPI.

```excel
=COUNTIF('21-10-2026 STK NPI.xls'!$A$4:$H$390;A5)
```

**Keterangan:**
- `$A$4:$H$390` merupakan rentang data referensi NPI.
- `A5` merupakan SKU yang diperiksa.
- Hasil `0` berarti SKU tidak ditemukan.
- Hasil lebih dari `0` berarti SKU ditemukan setidaknya satu kali.

## 🗓️ 3. Penggantian Nama File Berdasarkan Tanggal

Nama file referensi harus menggunakan tanggal stok hari sebelumnya.

Contoh:

| Tanggal Proses | Tanggal File Referensi |
|---|---|
| 22-10-2026 | `21-10-2026 STK AA.xls` |
| 22-10-2026 | `21-10-2026 STK NPI.xls` |

**Catatan:** Nama file pada rumus perlu diperbarui sesuai tanggal referensi yang digunakan. Rumus `COUNTIF` di atas menggunakan nama file tetap dan tidak otomatis mengganti tanggal.

## 🅱️ 4. Kondisi Bold untuk PT Produk

Gunakan rumus berikut pada fitur **Conditional Formatting** di Excel.

```excel
=NOT(ISNUMBER(SEARCH("-";$A5)))
```

**Cara menerapkan:**

1. Pilih rentang data yang ingin diformat.
2. Buka **Home → Conditional Formatting → New Rule**.
3. Pilih **Use a formula to determine which cells to format**.
4. Masukkan rumus di atas.
5. Klik **Format → Font → Bold**.
6. Klik **OK** untuk menyimpan aturan.

**Logika rumus:**
- Jika kolom A tidak mengandung tanda `-`, format bold diterapkan.
- Jika kolom A mengandung tanda `-`, format bold tidak diterapkan oleh aturan ini.

> Pastikan posisi baris pada rumus (`A5`) sesuai dengan baris pertama dari rentang yang dipilih.

## ✅ 5. Interpretasi Hasil

| Hasil COUNTIF | Arti |
|---:|---|
| `0` | SKU tidak ditemukan pada data referensi |
| `1` | SKU ditemukan satu kali |
| `2` atau lebih | SKU ditemukan lebih dari satu kali |

## 📝 Catatan

- Pastikan file referensi tersedia dan nama file sesuai dengan tanggal yang dibutuhkan.
- Pastikan rentang data mencakup seluruh SKU yang akan diperiksa.
- `COUNTIF` memeriksa kecocokan nilai, bukan jumlah stok.
- Jika SKU ditemukan beberapa kali, hasil menunjukkan jumlah kemunculan, bukan otomatis menandakan data valid.
- Aturan bold hanya mengatur format tampilan dan tidak mengubah nilai SKU.

---

**Dokumentasi:** SKU Stock Data Filter  
**Tools:** Microsoft Excel · Formula `COUNTIF` · Conditional Formatting
