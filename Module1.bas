Attribute VB_Name = "Module1"
'Contoh ini akan menampilkan lama waktu Windows telah 'berjalan dalam milidetik.
'1000 milidetik = 1 detik. Untuk mengkonversi dari 'milidetik ke detik, bagi hasilnya dengan 1000.

Declare Function GetTickCount Lib "kernel32" () As Long

