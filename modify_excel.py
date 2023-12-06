print("Script starting...")

import pandas as pd
from datetime import datetime
from difflib import SequenceMatcher


file_path = './sample.xlsx'
df = pd.read_excel(file_path)

print("\n\n")
print("DATA: ALL")
print(df)
print("\n\n")


print("-----------------------------------------START-------------------------------------------")
print("\n\n")


# ...............................................................................................
# validasi nomor polis endorsement


def process_nopolis(row):
    if 'X' in row['nopolis']:
        last_chars = row['nopolis'][-2:]
        if last_chars.isdigit():
            row['nopolis'] = row['nopolis'][:-2] + '{:02d}'.format(int(last_chars) + 1)
            slashes_count = row['blnbayar'].count('/')
            row['c01'] = -row['c01'] * (slashes_count - 12)
            row['blnbayar'] = '23/12'
            row['premi'] = row['c01']
            row['endorsement'] = True
        else:
            row['endorsement'] = False 
    else:
        row['endorsement'] = False 
    return row

df = df.apply(process_nopolis, axis=1)

# simpan baris polis endorsement baru
df_endorsement = df[df['endorsement']]

print("LIST POLIS ENDORSEMENT: df_endorsement")
print(df_endorsement)
print("\n\n")


# ...............................................................................................
# mengubah periodeawal pada polis endorsement baru


df_endorsement.loc[:, 'periodeawal'] = df_endorsement['periodeakhir'] - pd.DateOffset(years=1)

print("LIST POLIS ENDORSEMENT (perubahan periodeawal): df_endorsement")
print(df_endorsement)
print("\n\n")


# ...............................................................................................
# inisialisasi polis yang bukan endorsement


df_not_endorsement = df[~df['endorsement']]


# ...............................................................................................
# penghapusan polis induk yang sudah pernah diendorsement


def similarity_ratio(a, b):
    return SequenceMatcher(None, a, b).ratio()

def filter_similar_nopolis(row):
    for _, endorsement_row in df_endorsement.iterrows():
        
        # potong nomor polis sebelum karakter '.X'
        row_nopolis_cut = row['nopolis'].split('.X')[0]
        endorsement_nopolis_cut = endorsement_row['nopolis'].split('.X')[0]

        similarity = similarity_ratio(row_nopolis_cut, endorsement_nopolis_cut)
        if similarity == 1.0:  # jika nomor polis bentuknya sama hingga sebelum '.X'
            return False

    return True

df_not_endorsement = df_not_endorsement[df_not_endorsement.apply(filter_similar_nopolis, axis=1)].copy()

print("HASIL PENGHAPUSAN POLIS INDUK: df_not_endorsement")
print(df_not_endorsement)
print("\n\n")


# ...............................................................................................
# penghapusan polis klaim


df_not_endorsement = df_not_endorsement[df_not_endorsement['isklaim'] != 1].copy()

print("HASIL PENGHAPUSAN POLIS KLAIM: df_not_endorsement")
print(df_not_endorsement)
print("\n\n")


# .................................................................................................
# manipulasi polis yang bukan endorsement / mengendors polis yang belum pernah diendors


df_not_endorsement.loc[:, 'nopolis'] = df_not_endorsement['nopolis'] + '.X01'
df_not_endorsement.loc[:, 'c01'] = -df_not_endorsement['premi']
df_not_endorsement.loc[:, 'premi'] = -df_not_endorsement['premi']
df_not_endorsement.loc[:, 'saldopinjaman'] = -df_not_endorsement['saldopinjaman']
df_not_endorsement.loc[:, 'blnbayar'] = '23/12'
df_not_endorsement.loc[:, 'masa'] = df_not_endorsement['masa']
df_not_endorsement.loc[:, 'periodeawal'] = df_not_endorsement['periodeawal']
df_not_endorsement.loc[:, 'periodeakhir'] = df_not_endorsement['periodeakhir']

print("LIST POLIS YANG TIDAK ADA ENDORS: df_not_endorsement")
print(df_not_endorsement)
print("\n\n")


# .................................................................................................
# validasi periode akhir


current_date = datetime.now()
df_valid_tanggal = df_not_endorsement[df_not_endorsement['periodeakhir'] <= pd.to_datetime('2023-11-30')]

print("LIST POLIS YANG PERIODEAKHIR < 30 NOVEMBER 2023: df_valid_tanggal")
print(df_valid_tanggal)
print("\n\n")


# .................................................................................................
# gabungkan polis endorsement


df_result = pd.concat([df_valid_tanggal, df_endorsement])

print("DATA RESULT: df_valid_tanggal & df_endorsement")
print(df_result)
print("\n\n")


# .................................................................................................
# kirim ke file yang baru


output_file_path = './sample-result.xlsx'
df_result.to_excel(output_file_path, index=False)

print("Success...")