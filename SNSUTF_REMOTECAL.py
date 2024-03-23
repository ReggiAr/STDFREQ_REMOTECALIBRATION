#
# The MIT License (MIT)
#
# Copyright (c) 2024 Reggi Aryunadi
# 
# 
# This software is created for Remote Calibration services at
# the Time and Frequency Laboratory of SNSU-BSN (National Metrology Institute of Indonesia)
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
#
# Updated 23-03-2024 21:18 UTC(IDN)











# ---------- Libraries
from openpyxl import Workbook
import numpy as np
import matplotlib.pyplot as plt

print("="*10,"Laboratory of SNSU Time and Frequency","="*10)
print("="*10,"Time Standard Remote Calibration v.02","="*10,"\n")


while True:

    # client name for make excel file
    client_name = input("Client Name\t\t:")
    print("\n")
    workbook = Workbook()

    while True:

        # lists
        stdPrn = []
        stdSttime = []
        stdRefGPS = []
        sortedStdRefGPS =[]

        uutPrn = []
        uutSttime = []
        uutRefGPS = []
        sortedUutRefGPS =[]

        stdRefVal = []
        uutRefVal = []

        CstdRefVal = []
        CstdRefGPS = []
        CuutRefVal = []
        CuutRefGPS = []

        difValue = []

        # MJD name for make sheet name
        mjd = input ("Input MJD\t\t:")
        print("\n")
        worksheet = workbook.active
        worksheet.title = mjd

        # std input
        # std file using CGGTTS DATA FORMAT version 01
        stdFile = input("Standard File Name\t:")
        print("\n")

        # uut input
        # std file using CGGTTS DATA FORMAT can be 01 or 
        uutFile = input("UUT File Name\t\t:")
        uutFormat = input ("UUT Data Format\t\t:")
        print("\nStarting Calculating . . . \n")
        
        # ---------- Baca dan print ke excel data STD

        #Buka STD File
        with open(f'{stdFile}.txt', 'r') as Filestd:
        # Baca semua baris kecuali baris pertama
            stdLines = Filestd.readlines()

        nama_kolom = ['PRN','STTIME','REFGPS']
        header = stdLines[17].strip().split()
        std_indeks_kolom = [header.index(kolom) for kolom in nama_kolom]

        # Loop melalui setiap baris file
        for line in stdLines[19:]:  # Mulai dari baris kedelapan belas karena baris pertama adalah header
            data = line.strip().split()
    
            # Pastikan jumlah elemen dalam baris sesuai dengan jumlah kolom yang diharapkan
            if len(data) >= max(std_indeks_kolom) + 1:
                # Masukkan ke dalam list yang sesuai
                if data[std_indeks_kolom[0]].isdigit():
                    stdPrn.append(data[std_indeks_kolom[0]])
                else:
                    print("Skipping stdprn non-numeric values\n")

                if data[std_indeks_kolom[1]].isdigit():
                    stdSttime.append(data[std_indeks_kolom[1]])
                else:
                    print("Skipping stdSttime non-numeric values\n")

                if data[std_indeks_kolom[2]].isdigit() or data[std_indeks_kolom[2]].startswith("-"):
                    stdRefGPS.append(data[std_indeks_kolom[2]])
                else:
                    print("Skipping stdRefGPS non-numeric values\n")

        print ("Transfering STD data to Excel File . . .\n")
        cell = worksheet.cell(row=2,column=2)
        cell.value = "STD Data Unsorted"

        cell = worksheet.cell(row=3,column=2)
        cell.value = "STD PRN"
        for i, item in enumerate(stdPrn, start=4):
            cell = worksheet.cell(row=i, column=2)
            cell.value = item

        cell = worksheet.cell(row=3,column=3)
        cell.value = "STD STTIME"
        for i, item in enumerate(stdSttime, start=4):
            cell = worksheet.cell(row=i, column=3)
            cell.value = item

        cell = worksheet.cell(row=3,column=4)
        cell.value = "STD REFGPS"
        for i, item in enumerate(stdRefGPS, start=4):
            cell = worksheet.cell(row=i, column=4)
            cell.value = item

# ---------- Baca dan print ke excel data UUT

        #Buka STD File
        with open(f'{uutFile}.txt', 'r') as Fileuut:
        # Baca semua baris kecuali baris pertama
            uutLines = Fileuut.readlines()

        nama_kolom = ['PRN','STTIME','REFGPS']
        header = uutLines[17].strip().split()
        uut_indeks_kolom = [header.index(kolom) for kolom in nama_kolom]

        # Loop melalui setiap baris file
        for line in uutLines[19:]:  # Mulai dari baris kedelapan belas karena baris pertama adalah header
            data = line.strip().split()
    
            # Pastikan jumlah elemen dalam baris sesuai dengan jumlah kolom yang diharapkan
            if len(data) >= max(uut_indeks_kolom) + 1:
                # Masukkan ke dalam list yang sesuai
                if data[uut_indeks_kolom[0]].isdigit():
                    uutPrn.append(data[uut_indeks_kolom[0]])
                else:
                    print("Skipping uutPRN non-numeric values\n")
                if data[uut_indeks_kolom[1]].isdigit():
                    uutSttime.append(data[uut_indeks_kolom[1]])
                else:
                    print("Skipping uutSttime non-numeric values\n")
                if data[uut_indeks_kolom[2]].isdigit() or data[uut_indeks_kolom[2]].startswith("-"):
                    uutRefGPS.append(data[uut_indeks_kolom[2]])
                else:
                    print("Skipping uutRefGPS non-numeric values\n")

# - - - - - transfering data to excel
        print ("Transfering UUT data to Excel File . . .\n")
        cell = worksheet.cell(row=2,column=6)
        cell.value = "UUT Data Unsorted"

        cell = worksheet.cell(row=3,column=6)
        cell.value = "UUT PRN"
        for i, item in enumerate(uutPrn, start=4):
            cell = worksheet.cell(row=i, column=6)
            cell.value = item

        cell = worksheet.cell(row=3,column=7)
        cell.value = "UUT STTIME"
        for i, item in enumerate(uutSttime, start=4):
            cell = worksheet.cell(row=i, column=7)
            cell.value = item

        cell = worksheet.cell(row=3,column=8)
        cell.value = "UUT REFGPS"
        for i, item in enumerate(uutRefGPS, start=4):
            cell = worksheet.cell(row=i, column=8)
            cell.value = item

# - - - - - - - - - - - -  std  prn
        print ("Calculating Reference Value . . . \n")
        for i in range(len(stdPrn)):
            stdRefVal.append(int(stdPrn[i]) * int(stdSttime[i]))

        for i in range (len(uutPrn)):
            uutRefVal.append(int(uutPrn[i]) * int(uutSttime[i]))

        cell = worksheet.cell(row=2,column=10)
        cell.value = "Reference Value"
        cell = worksheet.cell(row=3,column=10)
        cell.value = "STD"
        for i, item in enumerate(stdRefVal, start=4):
            cell = worksheet.cell(row=i, column=10)
            cell.value = item
        cell = worksheet.cell(row=3,column=11)
        cell.value = "UUT"
        for i, item in enumerate(uutRefVal, start=4):
            cell = worksheet.cell(row=i, column=11)
            cell.value = item

# - - - - - - - Sorting Data 
        print("Sorting Data . . . . .\n")
        # Gabungkan stdRefVal dan stdRefGPS bersama-sama
        combined_std = list(zip(stdRefVal, stdRefGPS))
        combined_uut = list(zip(uutRefVal, uutRefGPS))

        # Urutkan berdasarkan stdRefVal
        sorted_combined_data_std = sorted(combined_std, key=lambda x: x[0])
        sorted_combined_data_uut = sorted(combined_uut, key=lambda x: x[0])

        # Pisahkan kembali sortedStdRefVal dan sortedStdRefGPS
        sortedStdRefVal = [data[0] for data in sorted_combined_data_std]
        sortedStdRefGPS = [data[1] for data in sorted_combined_data_std]
        sortedUutRefVal = [data[0] for data in sorted_combined_data_uut]
        sortedUutRefGPS = [data[1] for data in sorted_combined_data_uut]

        cell = worksheet.cell(row=2,column=13)
        cell.value = "Sorted Value"

        cell = worksheet.cell(row=3,column=13)
        cell.value = "STD Ref Val"
        for i, item in enumerate(sortedStdRefVal, start=4):
            cell = worksheet.cell(row=i, column=13)
            cell.value = item
        cell = worksheet.cell(row=3,column=14)
        cell.value = "STD RefGPS"
        for i, item in enumerate(sortedStdRefGPS, start=4):
            cell = worksheet.cell(row=i, column=14)
            cell.value = item

        cell = worksheet.cell(row=3,column=15)
        cell.value = "UUT Ref Val"
        for i, item in enumerate(sortedUutRefVal, start=4):
            cell = worksheet.cell(row=i, column=15)
            cell.value = item
        cell = worksheet.cell(row=3,column=16)
        cell.value = "UUT RefGPS"
        for i, item in enumerate(sortedUutRefGPS, start=4):
            cell = worksheet.cell(row=i, column=16)
            cell.value = item
# - - - - - - - - Matching Data
            
        print("Matching Data . . .\n")

        for a in sortedStdRefVal:
            for b in sortedUutRefVal:
                res = a/b
                if res == 1:
                    indexa = sortedStdRefVal.index(a)
                    indexb = sortedUutRefVal.index(b)
                    CstdRefVal.append(sortedStdRefVal[indexa])
                    CstdRefGPS.append(sortedStdRefGPS[indexa])
                    CuutRefVal.append(sortedUutRefVal[indexb])
                    CuutRefGPS.append(sortedUutRefGPS[indexb])
                    break
                else:
                    pass

        #print to excel
        cell = worksheet.cell(row=2,column=18)
        cell.value = "MATCH DATA"

        cell = worksheet.cell(row=3,column=18)
        cell.value = "STD REF VAL"

        for i, item in enumerate(CstdRefVal, start=4):
            cell = worksheet.cell(row=i, column=18)
            cell.value = item
        
        cell = worksheet.cell(row=3,column=19)
        cell.value = "STD RefGPS"
        for i, item in enumerate(CstdRefGPS, start=4):
            cell = worksheet.cell(row=i, column=19)
            cell.value = item
        cell = worksheet.cell(row=3,column=19)
        cell.value = "STD RefGPS"
        for i, item in enumerate(CstdRefGPS, start=4):
            cell = worksheet.cell(row=i, column=19)
            cell.value = item

        cell = worksheet.cell(row=3,column=21)
        cell.value = "UUT REF VAL"
        for i, item in enumerate(CuutRefVal, start=4):
            cell = worksheet.cell(row=i, column=21)
            cell.value = item
        cell = worksheet.cell(row=3,column=22)
        cell.value = "UUT REFGPS"
        for i, item in enumerate(CuutRefGPS, start=4):
            cell = worksheet.cell(row=i, column=22)
            cell.value = item

# calculating differences
        print("Calculating Differences . . .\n")

        for i in range (len(CstdRefGPS)):
            try:
                difValue.append(int(CstdRefGPS[i])-int(CuutRefGPS[i]))
            except:
                print("")
        

        cell = worksheet.cell(row=3,column=24)
        cell.value = "STD - DUT"
        for i, item in enumerate(difValue, start=4):
            cell = worksheet.cell(row=i, column=24)
            cell.value = item
        
        print("Calculating Average . . .\n")

        average = np.mean (difValue)
        cell = worksheet.cell(row=2,column=26)
        cell.value = "Average :"
        cell = worksheet.cell(row=3,column=26)
        cell.value = average

        print("Show and Save Graph :\n")

        # Membuat plot
        plt.plot(difValue)

        # Menambahkan judul dan label sumbu
        plt.title(f'Grafik Perbedaan Nilai Ref GPS UTC(IDN) dengan {client_name} pada {mjd}')
        plt.xlabel('Waktu')
        plt.ylabel(f'UTC(IDN) - {client_name}')

        plt.show()

        # Menyimpan plot ke dalam komputer
        plt.savefig(f'{client_name} - {mjd}.png')

        print("Calculating Done\n")
        print("Saving to Excel . . . .\n")
        # simpan excel
        excelFile = f"{client_name}.xlsx"
        workbook.save(excelFile)
        print(f"Data Saved in {client_name}.xlsx in sheet {mjd}\n")

        # Menutup sesi MJD
        mjdNext = input ("Calculate next MJD (y/n)\t:")
        if mjdNext == "n":
            break
    
    # Menutup sesi Client
    clientNext = input ("Calculate next Client (y/n)\t:")
    if clientNext == "n":
        break


print ("Thank you and See you in the next Remote Calibration")
    
    


