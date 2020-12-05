# API-REGION-HELLOWORLD

Membentuk kata "HELLO WORLD" dengan menggunakan bantuan Library API Visual Basic 6

<br>
<br>

## PENJELASAN
<p><b>Dengan menggunakan Lib "gdi32" API pada Visual Basic 6, anda dapat memanipulasi berbagai objek yang diinginkan. Function lib "gdi32" dapat diterapkan pada form, adapun fungsi-fungsi yang digunakan secara umum yaitu :</b><br>
a. PathToRegion<br>
Berfungsi untuk menciptakan suatu area dari jalur yang dipilih ke dalam area tertentu.<br>
b. PtlnRegion<br>
Berfungsi untuk menentukan apakah titik tertentu berada di dalam wilayah yang ditentukan.<br>
c. OffsetRegion<br>
Berfungsi untuk bergerak berada di area intern dengan batasan yang telah ditentukan.<br>
d. CreateRoundRectRgn<br>
Berfungsi untuk membuat form persegi panjang dengan sudut tumpul (bulat).<br>
e. CreateRectRgnIndirect<br>
Berfungsi untuk membuat area persegi panjang dari struktur RECT.<br>
f. CreateRectRgn<br>
Berfungsi untuk membentuk objek baru berbentuk persegi panjang.<br>
g. CreatePolyPolygonRgn<br>
Berfungsi untuk membuat area yang terdiri dari serangkaian poligon.<br>
h. CreatePolygonRgn<br>
Berfungsi untuk membentuk objek baru berbentuk polygon.<br>
i. CreateEllipticRgnIndirect<br>
Berfungsi untuk membuat area elips dari struktur RECT.<br>
j. CreateEllipticRgn<br>
Berfungsi untuk membentuk objek baru berbentuk elips atau bulat.<br>
k. CombineRgn<br>
Berfungsi untuk menggabungkan bagian yang berpotongan dari dua area yang berbeda.<br>
a. Angka 2 pada fungsi Combine bersifat fill object yang berarti mengisi objek pada bidang tertentu yang diinginkan.<br>
b. Angka 4 pada fungsi Combine bersifat remove object yang berarti menghilangkan objek pada bidang tertentu yang diinginkan.</p><br><br>
  
<p><b>Adapun fungsi tambahan lainnya :</b><br>
1. SetWindowRgn (handle, variabel, True)<br>
Handle di sini maksudnya adalah handle dari form ataupun kontrol lainnya yang akan diubah atau tak terkalahkan bentuknya untuk form yaitu Form.hwnd.<br>
2. Send Message untuk penampilan hasil output ke windows.<br>
3. ReleaseCapture untuk menimbulkan proses tombol mouse jadi bisa responsive.</p>

<br>
<br>

## DOKUMENTASI
<img src="https://user-images.githubusercontent.com/54527592/101258632-59e7a300-3756-11eb-9db8-4d554fe43307.jpg"/>
