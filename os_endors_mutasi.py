print("Script starting.....")

print("\n\n")
print("ENDORSEMENT POLIS MUTASI")
print("----------------------------")


import pandas as pd
from datetime import datetime

# input os_endors_mutasi
file_path = './sample/input.xlsx'

# output os_endors_mutasi result
output_file_path = './sample/output.xlsx'


df = pd.read_excel(file_path)
df['selisih'] = 0

print("\n\n")
print("DATA: ALL")
print("----------------")
print(df)
print("\n\n")

print("-----------------------------------------START-------------------------------------------")
print("\n\n")


# ...............................................................................................
# validasi nomor polis endorsement


for index, row in df.iterrows():
    if 'X' in row['nopolis']:
        base_nopolis = row['nopolis'].split('.X')[0]

        # mencari baris dengan nopolis yang mirip (tanpa '.X')
        base_row = df[df['nopolis'].str.contains(base_nopolis) & ~df['nopolis'].str.contains('.X')]

        # if baris mirip, hitung selisih bulan bayarnya
        if not base_row.empty:
            selisih = len(row['blnbayar'].split(',')) - len(base_row['blnbayar'].values[0].split(','))
            df.at[index, 'selisih'] = -selisih
            
            last_digit = int(row['nopolis'][-1]) + 1
            df.at[index, 'nopolis'] = row['nopolis'][:-1] + str(last_digit)
            df.at[index, 'c01'] = row['c01'] * -selisih
            df.at[index, 'periodeawal'] = base_row['periodeawal'].values[0]
            df.at[index, 'blnbayar'] = '23/12'
            df.at[index, 'premi'] = df.at[index, 'c01']

            # menghapus baris nopolis induk (yg sudah endorsement)
            df = df.drop(base_row.index)

print("LIST POLIS SETELAH MANIPULASI MUTASI ENDORSEMENT: df")
print("----------------------------------------------------")
print(df)
print("\n\n")


# .................................................................................................
# manipulasi polis yang bukan endorsement / mengendors polis yang belum pernah diendors


for index, row in df.iterrows():
    if 'X' not in row['nopolis']:
        df.at[index, 'nopolis'] = row['nopolis'] + '.X01'
        df.at[index, 'c01'] = -row['premi']
        df.at[index, 'premi'] = -row['premi']
        df.at[index, 'saldopinjaman'] = -row['saldopinjaman']
        df.at[index, 'blnbayar'] = '23/12'

print("LIST POLIS SETELAH MANIPULASI MUTASI YANG BELUM PERNAH ENDORSEMENT: df")
print("----------------------------------------------------------------------")
print(df)
print("\n\n")

# .................................................................................................
# validasi periode akhir


current_date = datetime.now()
df_valid_tanggal = df[df['periodeakhir'] <= pd.to_datetime('2023-11-30')]

print("LIST POLIS YANG PERIODEAKHIR < 30 NOVEMBER 2023: df_valid_tanggal")
print(df_valid_tanggal)
print("\n\n")


# .................................................................................................
# menyimpan hasil ke excel


df_valid_tanggal.to_excel(output_file_path, index=False)
print("Success.....")