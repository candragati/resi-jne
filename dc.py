import MSWinPrint
import sys
import string
from datetime import *
import win32con

# Identitas pengirim -------------------------------
company_name = "PT. BUMI DATAR"
alamat_pengirim = "Jl. Santai Hadiah Sepeda"
alamat_pengirim2 = "Kel. Pak Lurah, Kec. Pak Camat Cirebon"
# --------------------------------------------------


# Pilihan Kota ------------------------------------
def pilihKota():
	try:
		a={0:'Jakarta',1:'Tasik',2:'Bandung',3:'Tangerang'}
		for i in range(len(a)):
			print str(i)+" = "+a[i]

		receiver= input ("Masukkan alamat tujuan :")
	
		print "Tujuan : %s\n"%a[receiver]

		if receiver == 0:	
			pilihKota.company_name = "PT. GULANG GULING"
			pilihKota.addr = "Jl. Jalan Sore"
			pilihKota.addr2 = "Jakarta Utara"
		elif receiver == 1:
			pilihKota.company_name = "CV BANTING SETIR"
			pilihKota.addr = "Jl. Kaki Jauh Sekali"
			pilihKota.addr2 = "KOTA TASIKMALAYA"
		elif receiver == 2:
			pilihKota.company_name = "TOKO HURUHARA"
			pilihKota.addr = "Komplek Kost Kostan"
			pilihKota.addr2 = "Bandung 47000"
		elif receiver == 3:
			pilihKota.company_name = "KIOS KAOS"
			pilihKota.addr = "Perumahan Anti Tawuran"
			pilihKota.addr2 = "TANGERANG"
		else:
			print "\nPilihan tidak ada"
			pilihKota()
	except:
		print "\nPilihan tidak ada"
		pilihKota()
# --------------------------------------------------


# Pilihan Printer ----------------------------------
def pilihPrinter():
	b=MSWinPrint.listprinters()
	for i in range(len(b)):
		print [i], b[i]
	try:
		printer = input( "Printer yang dipilih:")
		print b[printer]
		pilihPrinter.x = b[printer]
		print "mencetak..."
	except:
		print "\nPilihan tidak ada"
		pilihPrinter()
# --------------------------------------------------



# pengaturan cetak ---------------------------------
def cetak():
	mm = 100 # satuan dalam milimeter, biar enak ngalinya
	doc = MSWinPrint.document(papersize="a4",printer=pilihPrinter.x,orientation = "portrait")
	doc.begin_document()
	doc.setfont("Lucida Console", 16,bold=0)

	doc.gambartext(company_name,(mm*30,mm*-22,mm*95,mm*-150),win32con.DT_LEFT)
	doc.gambartext(alamat_pengirim,(mm*13,mm*-27,mm*95,mm*-150),win32con.DT_LEFT)
	doc.gambartext(alamat_pengirim2,(mm*13,mm*-31,mm*95,mm*-150),win32con.DT_LEFT)

	doc.gambartext(pilihKota.company_name,(mm*112,mm*-22,mm*200,mm*-150),win32con.DT_LEFT)
	doc.gambartext(pilihKota.addr,(mm*97,mm*-27,mm*200,mm*-150),win32con.DT_LEFT)
	doc.gambartext(pilihKota.addr2,(mm*97,mm*-31,mm*200,mm*-150),win32con.DT_LEFT)


	tglSkrg = str(date.today())
	tgl= string.split(tglSkrg, '-')
	hari = tgl[2]
	bulan = tgl[1]
	tahun = tgl[0][2:]
	tanggal = hari+" "+bulan+" "+tahun

	doc.setfont("Lucida Console", 18,bold=0)
	doc.gambartext(tanggal,(mm*55,mm*-76,mm*200,mm*-150),win32con.DT_LEFT)

	doc.end_document()
	
# --------------------------------------------------

pilihKota()
pilihPrinter()
cetak()
sys.exit()