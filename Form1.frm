VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19200
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   ScaleHeight     =   9840
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_Load()
'===================================================================================================='
'BARIS KHUSUS'
'===================================================================================================='
'################################################################################'
'Parent'
bg = CreateRoundRectRgn(180, 300, 1100, 450, 20, 20) 'Background'
CombineRgn bg, bg, bg, 2

'################################################################################'
'Huruf H'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n1 = CreateRoundRectRgn(250, 320, 270, 430, 0, 0) 'Garis Vertikal Kiri'
n2 = CreateRoundRectRgn(250, 365, 320, 385, 0, 0) 'Garis Horizontal Tengah'
n3 = CreateRoundRectRgn(315, 320, 335, 430, 0, 0) 'Garis Vertikal Kanan'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n1, 4
CombineRgn bg, bg, n2, 4
CombineRgn bg, bg, n3, 4

'################################################################################'
'Huruf E'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n4 = CreateRoundRectRgn(350, 320, 370, 430, 0, 0) 'Garis Vertikal Kiri'
n5 = CreateRoundRectRgn(350, 320, 415, 340, 0, 0) 'Garis Horizontal Atas'
n6 = CreateRoundRectRgn(350, 365, 415, 385, 0, 0) 'Garis Horizontal Tengah'
n7 = CreateRoundRectRgn(350, 410, 415, 430, 0, 0) 'Garis Horizontal Bawah'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n4, 4
CombineRgn bg, bg, n5, 4
CombineRgn bg, bg, n6, 4
CombineRgn bg, bg, n7, 4

'################################################################################'
'Huruf L'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n8 = CreateRoundRectRgn(430, 320, 450, 430, 0, 0) 'Garis Vertikal Kiri'
n9 = CreateRoundRectRgn(430, 410, 480, 430, 0, 0) 'Garis Horizontal Bawah'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n8, 4
CombineRgn bg, bg, n9, 4

'################################################################################'
'Huruf L'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n10 = CreateRoundRectRgn(495, 320, 515, 430, 0, 0) 'Garis Vertikal Kiri'
n11 = CreateRoundRectRgn(495, 410, 545, 430, 0, 0) 'Garis Horizontal Bawah'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n10, 4
CombineRgn bg, bg, n11, 4

'################################################################################'
'Huruf O'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah")'
n12 = CreateEllipticRgn(555, 320, 620, 430) 'Lingkaran Luar'
n13 = CreateEllipticRgn(570, 335, 605, 415) 'Lingkaran Dalam'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n12, 4
CombineRgn bg, bg, n13, 2

'################################################################################'
'Huruf W'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n14 = CreateRoundRectRgn(690, 320, 730, 340, 0, 0) 'Garis Horizontal Atas Kiri'
n15 = CreateRoundRectRgn(710, 320, 730, 430, 0, 0) 'Garis Vertikal Kiri'
n16 = CreateRoundRectRgn(710, 410, 750, 430, 0, 0) 'Garis Horizontal Bawah Kiri'
n17 = CreateRoundRectRgn(740, 360, 760, 430, 0, 0) 'Garis Vertikal Tengah'
n18 = CreateRoundRectRgn(730, 410, 780, 430, 0, 0) 'Garis Horizontal Bawah Kanan'
n19 = CreateRoundRectRgn(770, 320, 790, 430, 0, 0) 'Garis Vertikal Kanan'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n14, 4
CombineRgn bg, bg, n15, 4
CombineRgn bg, bg, n16, 4
CombineRgn bg, bg, n17, 4
CombineRgn bg, bg, n18, 4
CombineRgn bg, bg, n19, 4

'################################################################################'
'Huruf O'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah")'
n20 = CreateEllipticRgn(800, 320, 860, 430) 'Lingkaran Luar'
n21 = CreateEllipticRgn(815, 335, 845, 415) 'Lingkaran Dalam'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n20, 4
CombineRgn bg, bg, n21, 2

'################################################################################'
'Huruf R'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n22 = CreateRoundRectRgn(860, 320, 900, 340, 0, 0) 'Garis Horizontal Atas Kiri'
n23 = CreateRoundRectRgn(870, 320, 890, 430, 0, 0) 'Garis Vertikal Kiri'
n24 = CreateRoundRectRgn(870, 320, 940, 380, 100, 1000) 'Objek Melingkar'
n25 = CreateEllipticRgn(890, 335, 900, 380) 'Lingkaran Dalam'
n26 = CreateRoundRectRgn(870, 360, 925, 390, 20, 20)  'Garis Horizontal Tengah'
n27 = CreateRoundRectRgn(900, 320, 920, 430, 0, 0)  'Garis Vertikal Kanan'
n28 = CreateRoundRectRgn(910, 415, 940, 430, 0, 0)  'Garis Horizontal Bawah Kanan'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n22, 4
CombineRgn bg, bg, n23, 4
CombineRgn bg, bg, n24, 4
CombineRgn bg, bg, n25, 2
CombineRgn bg, bg, n26, 4
CombineRgn bg, bg, n27, 4
CombineRgn bg, bg, n28, 4

'################################################################################'
'Huruf L'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n29 = CreateRoundRectRgn(950, 320, 970, 430, 0, 0) 'Garis Vertikal Kiri'
n30 = CreateRoundRectRgn(950, 410, 1000, 430, 0, 0) 'Garis Horizontal Bawah'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n29, 4
CombineRgn bg, bg, n30, 4

'################################################################################'
'Huruf D'
'("Perataan","Tinggi-Atas","Lebar","Tinggi-Bawah","Lebar-Elips","Tinggi-Elips")'
n31 = CreateRoundRectRgn(1005, 320, 1040, 340, 0, 0) 'Garis Horizontal Atas'
n32 = CreateRoundRectRgn(1005, 320, 1025, 430, 0, 0) 'Garis Vertikal Kiri'
n33 = CreateRoundRectRgn(1005, 410, 1040, 430, 0, 0) 'Garis Horizontal Bawah'
n34 = CreateRoundRectRgn(1005, 320, 1080, 430, 100, 1000) 'Objek Melingkar'
n35 = CreateRoundRectRgn(1030, 340, 1060, 410, 100, 1000) 'Objek Melingkar'
'"Parent","Parent","Sub","2=fill object/4=remove object"'
CombineRgn bg, bg, n31, 4
CombineRgn bg, bg, n32, 4
CombineRgn bg, bg, n33, 4
CombineRgn bg, bg, n34, 4
CombineRgn bg, bg, n35, 2



'View Hasil Combine'
SetWindowRgn Form1.hwnd, bg, True
End Sub

'Responsive Object -> Mouse Trigger'
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
    SendMessage Form1.hwnd, &HA1, 2, 0&
End Sub

