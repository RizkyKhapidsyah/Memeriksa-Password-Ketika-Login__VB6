VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Password ketika Login"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Hitung As Integer   'Deklarasi variabel global
Dim Jawab As String     'Deklarasi variabel global
Private Sub Form_Load()
  Hitung = 0  'Saat form diload, Hitung mula-mula masih
              '0
End Sub

Private Sub cmdOK_Click()
  'Ulangi selama text1 tdk sama dengan "Rahmat Putra"
  Do While Text1.Text <> "Puri Suantri"
    Jawab = Text1.Text = "Puri Suantri"  'Inisialisasi
                                         'Jawab
      'Jika Jawab tdk sama dengan "masino"
      If Jawab <> "masino" Then
         Hitung = Hitung + 1  'Counter bertambah satu
         Tampung (Hitung)     'Hitung ke fungsi Tampung
         If Hitung = 3 Then   'Jika Hitung = 3, maka...
            'Tampilkan pesan
            Print "Password Blocked!"
            Text1.Enabled = False   'Text1 tdk bisa
            'diakses
            cmdOK.Enabled = False   'cmdOK tdk bisa
            'diakses
            cmdCancel.Default = True 'Hanya cmdCancel
            'yg bisa diakses
         End If
         Exit Sub                'Keluar dari prosedur
      Else     'Jika Jawab = "Puri Suantri"
         Exit Do  'Keluar dari Loop
      End If
  Loop
  Print "Welcome"  'Tampilkan pesan sukses
  'terserah Anda.. setelah ini akan apa...?
End Sub

Function Tampung(Hitung)
Dim Hasil As Integer
    Hasil = 0  'Inisialisasi variabel Hasil
    Hasil = Hasil + Hitung
    Text1.SetFocus  'Fokuskan kursor ke text1 kembali
    SendKeys "{Home}+{End}"  'Highlight text1
    Print "Kesempatan ke-" & Hasil  'Tampilkan sudah
    'berapa kali salah password
End Function

Private Sub cmdCancel_Click()
   Unload Me     'Keluar dari program
End Sub

'Jika di text1 telah diisi, tombol OK siap dienter
Private Sub Text1_KeyPress(KeyAscii As Integer)
   cmdOK.Default = True
End Sub

