# -*- coding: utf-8 -*-
"""
Created on Mon Aug 13 09:16:45 2018

@author: Najib
"""

import os
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
import xlrd
import pyodbc
from datetime import datetime

#conn = pyodbc.connect('DRIVER={SQL Server};SERVER=52.187.121.85;PORT=1433;DATABASE=LeadIntelligenceTest;uid=sa;pwd=1Qaz2wsx3edc;Trusted_Connection=NO')
#c = conn.cursor()
#print("tes")
stripColumn = ['NoFaktur',
               'Kota',
               'KodePos',
               'Provinsi',
               'NoKTP',
               'NoHP']

tglColumn   = ['TanggalCetak',
               'TanggalMohon',
               'TanggalLahir']

blankColumn = ['FinanceCompany',
               'UangMuka',
               'TenorTahun',
               'Email',
               'JenisKelamin',
               'Agama',
               'Pekerjaan',
               'Pengeluaran',
               'Pendidikan',
               'PIC',
               'MerkSebelumnya',
               'JenisSebelumnya',
               'Fungsi',
               'Pemakai',
               'StatusRumah',
               'Facebook',
               'Twitter',
               'Instagram',
               'Youtube',
               'Hobi',
               'Remark']

selectColumn = ['No. Faktur',
                'KTP No',
                'No.HP',
                'Kode Kota',
                'Kode Dealer',
                'Kode Pos',
                'Kode Prov',
                'Tgl Cetak',
                'Tgl Mohon',
                'Tgl Lahir',
                'Nama',
                'No. Rangka',
                'Kode Mesin',
                'No. Mesin',
                'TIPE',
                'Warna',
                'Gender',
                'Alamat',
                'Kel',
                'Kec',
                'Email',
                'Cash Credit',
                'Finance Company',
                'Down Payment',
                'DP Aktual',
                'Cicilan',
                'Tenor Bulan',
                'Tenor Tahun',
                'Agama',
                'Pekerjaan',
                'Pengeluaran',
                'Pendidikan',
                'Status Handphone',
                'No. Telp',
                'No.HP',
                'Umur',
                'Range Umur',
                'diHubungi',
                '3JENIS',
                '6JENIS',
                'Merk',
                'Jenis',
                'Fungsi',
                'Pemakai',
                'SalesPerson',
                'Verify Date',
                'Status Verifikasi',
                'Hobi',
                'Facebook',
                'Twitter',
                'Instagram',
                'Youtube',
                'PIC',
                'Status Rumah',
                'Tipe ATPM',
                'Tipe Var Plus',
                'Cust No.',
                'RO Type',
                'RO Data',
                'Member No.',
                'REGION',
                'Remark',
                'Jenis Sales',                
                ]

df = pd.read_excel(r'C:\Users\Najib\Documents\AGIT\Customer Insight\Data\AHM\1Januari2018.xlsx')

df = df.rename(columns={'No. Faktur' : 'NoFaktur',
                        'KTP No' : 'NoKTP',
                        'No.HP' : 'NoHP',
                        'Kode Kota' : 'Kota',
                        'Kode Dealer' : 'KodeDealer',
                        'Kode Pos' : 'KodePos',
                        'Kode Prov' : 'Provinsi',
                        'Tgl Cetak' : 'TanggalCetak',
                        'Tgl Mohon' : 'TanggalMohon',
                        'Tgl Lahir' : 'TanggalLahir',
                        'No. Rangka' : 'NoRangka',
                        'Kode Mesin' : 'KodeMesin',
                        'No. Mesin' : 'NoMesin',
                        'TIPE' : 'TipeMotor',
                        'Gender' : 'JenisKelamin',
                        'Kel' : 'Kelurahan',
                        'Kec' : 'Kecamatan',
                        'Cash Credit' : 'CashCredit',
                        'Finance Company' : 'FinanceCompany',
                        'Down Payment' : 'UangMuka',
                        'DP Aktual' : 'UangMukaAktual',
                        'Tenor Bulan' : 'TenorBulan',
                        'Tenor Tahun' : 'TenorTahun',
                        'Status Handphone' : 'StatusHP',
                        'No. Telp' : 'NoTelp',
                        'No.HP' : 'NoHP',
                        'Range Umur' : 'RangeUmur',
                        'diHubungi' : 'KebersediaanDihubungi',
                        '3JENIS' : 'Jenis3',
                        '6JENIS' : 'Jenis6',
                        'Merk' : 'MerkSebelumnya',
                        'Jenis' : 'JenisSebelumnya',
                        'SalesPerson' : 'SalesPerson',
                        'Verify Date' : 'TanggalVerifikasi',
                        'Status Verifikasi' : 'StatusVerifikasi',
                        'Status Rumah' : 'StatusRumah',
                        'Tipe ATPM' : 'TipeATPM',
                        'Tipe Var Plus' : 'TipeVarPlus',
                        'Cust No.' : 'NoCustomer',
                        'RO Type' : 'ROType',
                        'RO Data' : 'ROData',
                        'Member No.' : 'NoAnggota',
                        'Jenis Sales' : 'JenisSales'})

for i in stripColumn:
#    data.append(df[i])
    df[i] = df[i].apply(lambda x: x.strip("'"))

for j in tglColumn:
    df[j] = df[j].apply(lambda x: x.strip("'"))
    df[j] = df[j].apply(lambda x: datetime.strptime(x, '%d%m%Y'))

for k in blankColumn:
#    print(k)
    df[k] = df[k].replace({'N' : np.NaN})

print(df.head(10))


#book = xlrd.open_workbook(r'C:\Users\Najib\Documents\AGIT\Customer Insight\Data\AHM\1Januari2018.xlsx')
#sheet = book.sheet_by_name('januari')

