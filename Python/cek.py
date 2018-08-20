import os
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
import xlrd
import pyodbc
from datetime import datetime

conn = pyodbc.connect('DRIVER={SQL Server};SERVER=52.187.121.85;PORT=1433;DATABASE=LeadIntelligenceTest;uid=sa;pwd=1Qaz2wsx3edc;Trusted_Connection=NO')
c = conn.cursor()
#print("tes")

df = pd.read_excel(r'C:\Users\Najib\Documents\AGIT\Customer Insight\Data\AHM\1Januari2018.xlsx')

stripColumn = ['No. Faktur',
               'Kode Kota',
               'Kode Pos',
               'Kode Prov',
               'KTP No',
               'No.HP']

tglColumn   = ['Tgl Cetak',
               'Tgl Mohon',
               'Tgl Lahir']

blankColumn = ['Finance Company',
               'Down Payment',
               'Tenor Tahun',
               'Email',
               'Gender',
               'Agama',
               'Pekerjaan',
               'Pengeluaran',
               'Pendidikan',
               'PIC',
               'Merk',
               'Jenis',
               'Fungsi',
               'Pemakai',
               'Status Rumah',
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
#print(df['Agama'])
#data.replace({'very bad': 1, 'bad': 2, 'poor': 3, 'good': 4, 'very good': 5}, inplace=True)
#print (df.replace({{np.NaN : 'N'}))
##print(df['Agama'])
#print (df.replace({'Flag Hobi' : {np.NaN : 'N'}}))

# Open the workbook and define the worksheet
book = xlrd.open_workbook(r'C:\Users\Najib\Documents\AGIT\Customer Insight\Data\AHM\1Januari2018.xlsx')
sheet = book.sheet_by_name('januari')

#print(book['No. Faktur'][r].strip("'"))

#print(sheet.nrows)

#for r in range(1, sheet.nrows):
#    
#noFaktur = df['No. Faktur']
#print(noFaktur)
#data = []
#
#
#newData = []
#
for i in stripColumn:
#    data.append(df[i])
    df[i] = df[i].apply(lambda x: x.strip("'"))

for j in tglColumn:
    df[j] = df[j].apply(lambda x: x.strip("'"))
    df[j] = df[j].apply(lambda x: datetime.strptime(x, '%d%m%Y'))

for k in blankColumn:
#    print(k)
    df[k] = df[k].replace({'N' : np.NaN})

print(df['Agama'])
#print(data[5])
    
#data = df[selectColumn]
#print(data)
#print(data['No. Faktur'])

#query = "INSERT INTO TB_R_SALES_ORDER (DealerID,CustomerID,KodeDealer,Nama,NoKTP,TanggalCetak,TanggalMohon,NoFaktur,NoRangka,KodeMesin,NoMesin,TipeMotor,Warna,JenisKelamin,TanggalLahir,Alamat,Kelurahan,Kecamatan,Kota,KodePos,Provinsi,Email,CashCredit,FinanceCompany,UangMuka,UangMukaAktual,Cicilan,TenorBulan,TenorTahun,Agama,Pekerjaan,Pengeluaran,Pendidikan,StatusHP,NoTelp,NoHP,Umur,RangeUmur,KebersediaanDihubungi,Jenis3,Jenis6,MerkSebelumnya,JenisSebelumnya,Fungsi,Pemakai,SalesPerson,TanggalVerifikasi,StatusVerifikasi,Hobi,Facebook,Twitter,Instagram,Youtube,PIC,StatusRumah,TipeATPM,TipeVarPlus,NoCustomer,ROType,ROData,NoAnggota,Region,Remark,JenisSales,RowStatus,CreatedBy,CreatedDate) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
#
##for r in data:
#DealerID = '5',
#CustomerID = data['Nama'],
#KodeDealer = data['Kode Dealer'],
#Nama = data['Nama'],
#NoKTP = data['KTP No'],
#TanggalCetak = data['Tgl Cetak'],
#TanggalMohon = data['Tgl Mohon'],
#NoFaktur = data['No. Faktur'],
#NoRangka = data['No. Rangka'],
#KodeMesin = data['Kode Mesin'],
#NoMesin = data['No. Mesin'],
#TipeMotor = data['TIPE'],
#Warna = data['Warna'],
#JenisKelamin = data['Gender'],
#TanggalLahir = data['Tgl Lahir'],
#Alamat = data['Alamat'],
#Kelurahan = data['Kel'],
#Kecamatan = data['Kec'],
#Kota = data['Kode Kota'],
#KodePos = data['Kode Pos'],
#Provinsi = data['Kode Prov'],
#Email = data['Email'],
#CashCredit = data['Cash Credit'],
#FinanceCompany = data['Finance Company'],
#UangMuka = data['Down Payment'],
#UangMukaAktual = data['DP Aktual'],
#Cicilan = data['Cicilan'],
#TenorBulan = data['Tenor Bulan'],
#TenorTahun = data['Tenor Tahun'],
#Agama = data['Agama'],
#Pekerjaan = data['Pekerjaan'],
#Pengeluaran = data['Pengeluaran'],
#Pendidikan = data['Pendidikan'],
#StatusHP = data['Status Handphone'],
#NoTelp = data['No. Telp'],
#NoHP = data['No.HP'],
#Umur = data['Umur'],
#RangeUmur = data['Range Umur'],
#KebersediaanDihubungi = data['diHubungi'],
#Jenis3 = data['3JENIS'],
#Jenis6 = data['6JENIS'],
#MerkSebelumnya = data['Merk'],
#JenisSebelumnya = data['Jenis'],
#Fungsi = data['Fungsi'],
#Pemakai = data['Pemakai'],
#SalesPerson = data['SalesPerson'],
#TanggalVerifikasi = data['Verify Date'],
#StatusVerifikasi = data['Status Verifikasi'],
#Hobi = data['Hobi'],
#Facebook = data['Facebook'],
#Twitter = data['Twitter'],
#Instagram = data['Instagram'],
#Youtube = data['Youtube'],
#PIC = data['PIC'],
#StatusRumah = data['Status Rumah'],
#TipeATPM = data['Tipe ATPM'],
#TipeVarPlus = data['Tipe Var Plus'],
#NoCustomer = data['Cust No.'],
#ROType = data['RO Type'],
#ROData = data['RO Data'],
#NoAnggota = data['Member No.'],
#Region = data['REGION'],
#Remark = data['Remark'],
#JenisSales = data['Jenis Sales'],
#RowStatus = 1,
#CreatedBy = 'DSSCI_5',
#CreatedDate = datetime.now()
#
#values = (DealerID,CustomerID,KodeDealer,Nama,NoKTP,TanggalCetak,TanggalMohon,NoFaktur,NoRangka,KodeMesin,NoMesin,TipeMotor,Warna,JenisKelamin,TanggalLahir,Alamat,Kelurahan,Kecamatan,Kota,KodePos,Provinsi,Email,CashCredit,FinanceCompany,UangMuka,UangMukaAktual,Cicilan,TenorBulan,TenorTahun,Agama,Pekerjaan,Pengeluaran,Pendidikan,StatusHP,NoTelp,NoHP,Umur,RangeUmur,KebersediaanDihubungi,Jenis3,Jenis6,MerkSebelumnya,JenisSebelumnya,Fungsi,Pemakai,SalesPerson,TanggalVerifikasi,StatusVerifikasi,Hobi,Facebook,Twitter,Instagram,Youtube,PIC,StatusRumah,TipeATPM,TipeVarPlus,NoCustomer,ROType,ROData,NoAnggota,Region,Remark,JenisSales,RowStatus,CreatedBy,CreatedDate)
#
##print(KodeDealer)    
#c.execute(query, values)
##
#conn.commit()
##
#c.close()
##
#conn.close()
#
#print('All Done')

#for i in df:
#    print(i)


#data_new = data[:802]
#
#print(data[2600:])
#
#number_rows_select = 100
#number_loop = int(len(data_new)/number_rows_select)
#sisa = len(data_new)%number_rows_select
#
#for i in range(0,number_loop):
#    print('=========== Batch {} ==========='.format(i))
#    print(data_new[i*number_rows_select:(i+1)*number_rows_select])
#    print('\n\n\n')
#    insert ke db
#
#if sisa > 0:
#    print(data_new[number_loop*number_rows_select:])
#    insert ke db
    
#print(len(df))
#print(number_loop)
#print(sisa)

#if sisa > 0 :
#    insert  df[number_loop*number_rows_select:]
    
    
    
    
#for i range(0,number_loop):
#    df[i*100:(i+1)*100]
    
#    noFaktur = df['No. Faktur'].apply(lambda x: x.strip("'"))
#    print(noFaktur)
#    noRangka    = sheet.cell(r,1).value
#    kodeMesin   = sheet.cell(r,2).value
#    noMesin     = sheet.cell(r,3).value
#    ctk = df['Tgl Cetak'].apply(lambda x: x.strip("'"))
#    tglCetak   = ctk.apply(lambda x: datetime.strptime(x, '%d%m%Y'))
#    mhn = df['Tgl Mohon'].apply(lambda x: x.strip("'"))
#    tglMohon   = mhn.apply(lambda x: datetime.strptime(x, '%d%m%Y'))
#    lhr = df['Tgl Lahir'].apply(lambda x: x.strip("'"))
#    tglLahir   = lhr.apply(lambda x: datetime.strptime(x, '%d%m%Y'))
#    nama        = sheet.cell(r,6).value
#    alamat      = sheet.cell(r,7).value
#    kel         = sheet.cell(r,8).value
#    kec         = sheet.cell(r,9).value
#    kota        = df['Kode Kota'].apply(lambda x: x.strip("'"))
#    kodePos     = df['Kode Pos'].apply(lambda x: x.strip("'"))
#    provinsi    = df['Provinsi'].apply(lambda x: x.strip("'"))
#    cashCredit  = sheet.cell(r,13).value
#    kodeDealer  = sheet.cell(r,14).value
#    noKTP       = df['KTP No'][r].strip("'")
#    print(noKTP)
#    financeComp = sheet.cell(r,16).value
#    uangMuka    = sheet.cell(r,17).value
#    tenorTahun  = sheet.cell(r,18).value
#    email       = sheet.cell(r,19).value
#    jenisSales  = sheet.cell(r,23).value    
#    gender      = sheet.cell(r,24).value
#    agama       = sheet.cell(r,32).value
#    pekerjaan   = sheet.cell(r,33).value
#    pengeluaran = sheet.cell(r,34).value
#    pendidikan  = sheet.cell(r,35).value
#    pic         = sheet.cell(r,36).value
#    telp        = sheet.cell(r,38).value
#    dihubungi   = sheet.cell(r,39).value
#    merk        = sheet.cell(r,40).value
#    jenis       = sheet.cell(r,41).value
#    pemakai     = sheet.cell(r,43).value
#    salesPerson = sheet.cell(r,44).value
#    umur        = sheet.cell(r,45).value
#    rangeUmur   = sheet.cell(r,46).value
#    region      = sheet.cell(r,47).value
#    tipe        = sheet.cell(r,48).value
#    jenis6      = sheet.cell(r,49).value
#    jenis3      = sheet.cell(r,50).value
#    statusRumah = sheet.cell(r,64).value
#    statusHP    = sheet.cell(r,65).value
#    statusVerif = sheet.cell(r,66).value
#    dpAktual    = sheet.cell(r,67).value
#    tenorBulan  = sheet.cell(r,68).value
#    cicilan     = sheet.cell(r,69).value
#    tipeATPM    = sheet.cell(r,70).value
#    warna       = sheet.cell(r,71).value
#    tipeVarPlus = sheet.cell(r,72).value
#    noCustomer  = sheet.cell(r,73).value
#    roType      = sheet.cell(r,74).value
#    roData      = sheet.cell(r,75).value
#    noMember    = sheet.cell(r,76).value
#    facebook    = sheet.cell(r,77).value
#    twitter     = sheet.cell(r,78).value
#    instagram   = sheet.cell(r,79).value
#    youtube     = sheet.cell(r,80).value
#    hobi        = sheet.cell(r,81).value
#    remark      = sheet.cell(r,82).value
#
#tgl = df['Tgl Cetak'].apply(lambda x: x.strip("'"))
#print(tgl.apply(lambda x: datetime.strptime(x, '%d%m%Y')))
#    
#for columname in stripcolumn:
##   print (columname)
#    print(df[columname].apply(lambda x: x.strip("'")))
#    
#for column in tglcolumn:
    

#print(df['No. Faktur'][0].strip("'"))
#    a = datetime.strptime(df['Tgl Cetak'][r].strip("'"), '%d%m%Y')
#    print(a)
