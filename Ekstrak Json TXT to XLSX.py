import json
import pandas as pd

try:
    with open('outputinvoice_list.txt', 'r', encoding='utf-8') as f:
        data = json.load(f)['Payload']['Data']
    
    df_raw = pd.json_normalize(data)
    
    mapping = {
        'BuyerTIN': 'NPWP Pembeli / Identitas lainnya',
        'BuyerTaxpayerNameClear': 'Nama Pembeli',
        'TaxInvoiceNumber': 'Nomor Faktur Pajak',
        'TaxInvoiceDate': 'Tanggal Faktur Pajak',
        'TaxInvoiceYear': 'Tahun',
        'TaxInvoiceStatus': 'Status Faktur',
        'SellingPrice': 'Harga Jual/Penggantian/DPP',
        'OtherTaxBase': 'DPP Nilai Lain/DPP',
        'VAT': 'PPN',
        'STLG': 'PPnBM',
        'Signer': 'Penandatangan',
        'ReportedByBuyer': 'Dilaporkan oleh Pembeli',
        'ReportedBySeller': 'Dilaporkan oleh Penjual',
        'ReportedByVATCollector': 'Dilaporkan oleh Pemungut PPN',
        'LastUpdatedDate': 'Terakhir Diperbarui',
        'CreationDate': 'Tanggal Dibuat',
        'YearCredit': 'Tahun Kredit',
        'DocumentFormNumber': 'Nomor Form Dokumen',
        'Reference': 'Referensi',
        'DisplayName': 'Nama Tampilan',
        'SellerLastUpdatedDate': 'Terakhir Diperbarui Oleh Penjual'
    }
    
    cols = [c for c in mapping.keys() if c in df_raw.columns]
    df_clean = df_raw[cols].rename(columns=mapping)
    
    with pd.ExcelWriter('Hasil Ekstrak Data Json.xlsx', engine='xlsxwriter') as writer:
        df_clean.to_excel(writer, sheet_name='Hasil Ekstrak Bersih', index=False)
        df_raw.to_excel(writer, sheet_name='Ekstrak Data', index=False)
        
        for sheet, df_sheet in [('Hasil Ekstrak Bersih', df_clean), ('Ekstrak Data', df_raw)]:
            worksheet = writer.sheets[sheet]
            for idx, col in enumerate(df_sheet.columns):
                content_max_len = 0
                if not df_sheet.empty:
                    content_max_len = df_sheet[col].astype(str).str.len().max()
                    if pd.isna(content_max_len):
                        content_max_len = 0
                
                max_len = max(content_max_len, len(str(col))) + 2
                worksheet.set_column(idx, idx, max_len)
    
    print("--> Proses konversi selesai file excel telah dibuat dengan nama Hasil Ekstrak Data Json.xlsx")
except Exception as e:
    print(f"--> Proses gagal: {e}")

input("--> Tekan enter untuk keluar")