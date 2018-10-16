VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAna 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10215
   ForeColor       =   &H80000018&
   Icon            =   "frmYonetim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   10215
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   405
      TabIndex        =   8
      Top             =   120
      Width           =   1635
   End
   Begin VB.CommandButton cmdCalistir 
      Caption         =   "�ALI�TIR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3645
      TabIndex        =   5
      Top             =   360
      Width           =   2040
   End
   Begin VB.ListBox lstSipler 
      Height          =   3375
      Left            =   45
      TabIndex        =   4
      Top             =   1665
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   5085
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ListBox lstDurum 
      Height          =   3375
      Left            =   2250
      TabIndex        =   2
      Top             =   1665
      Width           =   7890
   End
   Begin VB.ComboBox cboZaman 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmYonetim.frx":000C
      Left            =   9540
      List            =   "frmYonetim.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   405
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9720
      Top             =   765
   End
   Begin VB.Label lblBilgi 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2250
      TabIndex        =   7
      Top             =   1350
      Width           =   7035
   End
   Begin VB.Label lblListe 
      Height          =   195
      Left            =   45
      TabIndex        =   6
      Top             =   1440
      Width           =   1725
   End
   Begin VB.Label lblZaman 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9390
      TabIndex        =   0
      Top             =   135
      Width           =   750
   End
End
Attribute VB_Name = "frmAna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Dim sipNo As String
Dim sipListe() As String

Dim renk As String
Dim kullanici As String
Dim fso, fso1, fso2 As FileSystemObject
Dim f, f1 As Folder, sf, sf1 As Folder, fl, fl1, fl2, fl3, tf, tf1 As File, path As String



'Timer de�i�kenleri
Dim san, dak, i, y As Integer


'VERS�YONLAR
'v1.0.1:    1. Faz g�ncellemeleri bitti
'v1.0.2:    Versiyon sistemine ge�ildi
'v1.0.3:    Log sistemi kuruldu.
'v1.0.3:    Hatalar giderildi.
'v1.0.4:    Sadece txt listeleri i�liyor.
'v1.0.8:    Hatalar loglan�yor.
'v1.0.48:   Art�k sipari�ler csvden okunuyor
'v2.0.0:    Art�k durumlar� sunucudan okuyor.
'v2.0.9:    Art�k pdfleri renklerine g�re klas�re al�yor.
'v3.0.0:    Art�k t�m Grafikerler i�in kullan�labilir. Parametrik oldu
'v3.0.1:    Loglama sistemi g�ncellendi.
'v3.0.10:   �al��ma klas�r�nde Pdf yoksa klas�r� ta��ma i�lemini yapm�yor
'v3.0.96:   Kopyalama hatas�n� giderdim. Fso ile de�il filecopy ile kopyalama yap�yor


Private Sub cmdCalistir_Click()
'Yeni sipari� klas�rleri olu�turulur
salih_yeniler

'Sipari�ler listesini temizle
lstSipler.Clear

'Butonu kilitle
cmdCalistir.Caption = "�ALI�IYOR"
cmdCalistir.Enabled = False

Set fso1 = New FileSystemObject

'�nternetten biten sipari�leri listeye al�p diziye al�yor.
Call grafik_bitenler

'Listeye veri gelmezse ��k
If lstSipler.ListCount < 1 Then GoTo bitir:

'�al��acak grafikerleri bulal�m
tmpGrafikerler = Split(dict.Item("grafikerler"), ",")

'HER GRAF�KERDE D�NEL�M
For Each grafikerNo In tmpGrafikerler
'On Error Resume Next
DoEvents


    'Grafiker ad� belirleniyor.
    kullanici = dict.Item("grafikerAdi" & grafikerNo)
    'lblBilgi.Caption = kullanici
  
    
    'Grafiker logunu ekleyelim
    'Call log_ekle("Bilgi", kullanici, " ba�lad�")

    'Temp yollar belirleniyor
    tmpDevamKlasor = dict.Item("devamKlasor" & grafikerNo)
    tmpBitenKlasor = dict.Item("bitenKlasor" & grafikerNo)
    tmpPdfKlasor = dict.Item("pdfKlasor" & grafikerNo)
    
    Set f1 = fso1.GetFolder(tmpDevamKlasor)
    pbar.Max = f1.SubFolders.Count + 1
    
    'De�i�kenler
    sayac = 0
    sf_say = 0

    'Devam klas�rlerinde gezinti yapal�m.
    For Each sf1 In f1.SubFolders
    On Error Resume Next
    DoEvents
    
        'E�er klas�r bo� gelirse ge�
        If sf1.Name = "" Then GoTo klasor_gec:
        
        'De�i�kenler
        sf_say = sf_say + 1
        pbar.Value = sf_say
        tmpRenk = 0
        tmpBoyut = ""
        
        'Gezerken kullan�c� ve klas�r ad�n� g�sterelim
        lblBilgi.Caption = kullanici & ":" & sf1.Name
       
        'Klas�rden sipari� numaras�n� alal�m
        sipNo = Left(sf1.Name, 9)
              
        'Sipari� klas�r� gelen sipari� dizisi i�erisinde varsa
        If InStr(Join(sipListe), sipNo) > 0 Then
        
            'Sipari�in bilgilerini bulal�m
            For y = 0 To lstSipler.ListCount - 1
            
                tmpVeri = Split(lstSipler.List(y), "|")
                
                If tmpVeri(0) = sipNo Then
                'MsgBox sipNo
                    'renk
                    tmpRenk = tmpVeri(1)
                    If tmpVeri(1) = "" Then tmpRenk = "0"
                    
                    'Grafiker
                    tmpGrafiker = tmpVeri(2)
                    
                    'Miktar
                    tmpMiktar = tmpVeri(3)
                    If tmpVeri(3) = "" Then tmpMiktar = 0
                    
                    'Boyut
                    tmpBoyut = tmpVeri(4)
                    If tmpVeri(4) = "" Then tmpBoyut = ""

                    'Listede ilk buldu�u veri ile i�lem yap�p d�ng�den ��k�yor
                    Exit For
                End If
            Next y
            
            'Daha �nce biten klas�r�ne ta��nm��sa ge�
            If Dir(tmpBitenKlasor & "\" & sf1.Name, vbDirectory) <> "" Then
                Call log_ekle("Hata", kullanici, sipNo & " �nceden ta��nm��")
                'lstSipler.Text = sipNo & "|" & renk
                'lstSipler.List(lstSipler.ListIndex) = "x" & lstSipler.List(lstSipler.ListIndex)
                GoTo klasor_gec:
            End If
    
             'Klas�rleri belirleyelim
            kaynakYol = tmpDevamKlasor & "\" & sf1.Name
            hedefYol = tmpBitenKlasor & "\" & sf1.Name
            pdfkaynak = kaynakYol & "\" & sipNo & ".pdf"
            
            
            
            
            '*****Klas�r�n Pdf dosyas� varsa i�lem yap yoksa ge�
            If Dir(kaynakYol & "\" & sipNo & ".pdf") <> "" Then
            'MsgBox "Klas�rde PDF var bitenlere ta��ma i�lemi yap�yor"
            
                If Dir(tmpPdfKlasor & "\" & renk & "\" & sipNo & ".pdf") <> "" Then
                    Set fl1 = fso1.CreateTextFile(tmpPdfKlasor & "\" & sipNo & " PDF_onceden_tasinmis_uzerine_yazilacak", True)
                    Set fl1 = Nothing
                    
                    Call log_ekle("Bilgi", kullanici, sipNo & " Pdf �nceden ta��nm��. �zerine yaz�lacak")
                    'lstDurum.AddItem Now & " " & sipNo & " isimli PDF �nceden al�nm��"
                End If
                
              
                'Pdfyi ta��yal�m
                pdfhedef = tmpPdfKlasor & "\" & tmpRenk & " Renk\" & sipNo & "_" & tmpBoyut & "_" & tmpMiktar & ".pdf"
                FileCopy pdfkaynak, pdfhedef
                
                
                '1 sn bekle
                Sleep 1000
                
                'PDF ta��nm��sa klas�r� de ta��yal�m
                If Dir(pdfhedef) <> "" Then
                    'Klas�r� tamamlanan b�l�m�ne alal�m
                    fso1.MoveFolder kaynakYol, hedefYol
                    
                    '3 sn bekle
                    Sleep 3000
                
                    'Klas�r ta��namaz ise uyar� verelim.
                    If Dir(hedefYol, vbDirectory) = "" Then
                        'Hatay� kullan�c�n�n PDF klas�r�ne yazal�m
                        Set fl1 = fso1.CreateTextFile(tmpPdfKlasor & "\" & sipNo & " Klasoru_tasinamadi", True)
                        Set fl1 = Nothing
                        Call log_ekle("Hata", kullanici, sipNo & " Klasoru_tasinamadi")
                       
                    Else
                        'Ta��nan klas�r sayac�n� artt�ral�m
                        sayac = sayac + 1
                    End If
                    
                End If
                
            Else
                'Hatay� kullan�c�n�n PDF klas�r�ne yazal�m
                Set fl2 = fso1.CreateTextFile(tmpPdfKlasor & "\" & sipNo & " PDFsi_yok_yada_tasinamadi", True)
                Set fl2 = Nothing
                'Loga ekleyelim
                Call log_ekle("Hata", kullanici, sipNo & " PDFsi_yok_yada_tasinamadi")
                
            End If
            
        End If
    
    'Hata olursa yaz
    If Err Then
        'Hatay� kullan�c�n�n PDF klas�r�ne yazal�m
        Set fl3 = fso1.CreateTextFile(tmpPdfKlasor & "\" & sipNo & " " & Err.Description, True)
        
        Set fl3 = Nothing
        Call log_ekle("Hata", kullanici, lblBilgi.Caption & "*" & Err.Description)
        Err.Clear
    End If

    
klasor_gec:
    Next 'Devam klas�r�ndeki gezinti
        
    'Log atal�m
    If sayac > 0 Then
        Call log_ekle("Bilgi", kullanici, sayac & " sipari� ta��nd�")
    End If
    
    'Devam eden klas�r� temizleyelim
    Set f1 = Nothing

    If Err Then
        Call log_ekle("Hata", kullanici, Err.Description)
        Err.Clear
    End If
    
Next
    
'FSO kapat

Set fso1 = Nothing


'SIFIRLA
Call sifirla
    


Exit Sub
'Hata durumunda
bitir:
Call log_ekle("Hata", "Genel", lblBilgi.Caption & "***" & Err.Description)
Call sifirla 'S�f�rlama yapal�m
End Sub

Sub sifirla()
pbar.Value = 0
cmdCalistir.Caption = "�ALI�TIR"
cmdCalistir.Enabled = True
Erase sipListe()
lblBilgi.Caption = ""
lblBilgi.Caption = ""

End Sub

Sub grafik_bitenler()
Call baglan

SQL = "SELECT DISTINCT " _
& "siparis.siparis_no,siparis.renk_adet,siparisdurum.atanan_kullanici_id,siparis.toplam_adet,siparis.karton_ebat,siparis.son_durum_id,siparis.kullanici_id,siparis.yurt_id," _
& "siparisdurumtanim.adi " _
& "FROM siparisdurumtanim " _
& "INNER JOIN siparis ON siparis.son_durum_id = siparisdurumtanim.id " _
& "INNER JOIN siparisdurum ON siparisdurum.siparis_id = siparis.id " _
& "WHERE siparis.siparis_no like '19%' AND siparisdurum.atanan_kullanici_id IN (" & dict.Item("grafikerler") & ") AND siparis.son_durum_id IN (" & dict.Item("durumlar") & ")" _
& "ORDER BY siparisdurum.atanan_kullanici_id ASC"
'Debug.Print SQL
rs.Open SQL, con, 3, 4


i = 0
Do Until rs.EOF = True
    i = i + 1
    lstSipler.AddItem rs(0) & "|" & rs(1) & "|" & rs(2) & "|" & rs(3) & "|" & rs(4)
    rs.MoveNext
Loop

lblListe.Caption = "Toplam: " & i

'ba�lant�y� kapatal�m
Call kapat

'Sipari� listesini diziye atal�m
ReDim sipListe(lstSipler.ListCount)
For i = 0 To lstSipler.ListCount - 1
    
    tmpSipNo = Split(lstSipler.List(i), "|")
    sipListe(i) = tmpSipNo(0)
    
Next i

End Sub


Sub salih_yeniler()
On Error Resume Next
'Grafikere atanan sipari�ler
Call grafik_atanan

For i = 0 To lstSipler.ListCount - 1
    
    tmpSatir = Split(lstSipler.List(i), "|")
    sipNo = tmpSatir(0)
    tmpFirma = tmpSatir(1)
    SipTur = tmpSatir(2)
    
    If tmpSatir(2) = 1 Then SipTur = ""
    If tmpSatir(2) = 2 Then SipTur = "__DGS"
    If tmpSatir(2) = 3 Then SipTur = "__GSA"
    
    firma = CStr(tmpSatir(1))
    firma = fncYazimDuzeni(firma)
    bosluk = InStr(1, firma, " ", vbTextCompare)
    If bosluk > 0 Then firma = Left(firma, bosluk)
    firma = Trim(firma)
    kaynak = dict.Item("pdfKlasor12") & "\Sablon\Sablon.ai"
    hedef = dict.Item("kokKlasor12") & "\" & sipNo & " " & firma & SipTur & "\"
    
    If Dir(hedef, vbDirectory) = "" Then
        MkDir hedef
        FileCopy kaynak, hedef & sipNo & ".ai"
    End If
    
Next i

End Sub
Public Function fncYazimDuzeni(param1)
Dim aranacak, degistir As String

param1 = UCase(Trim(param1))
'param = Replace(param, "�", "I")

aranacak = "�,�,�,�,�,�,�"
degistir = "C,G,S,O,U,I,I"
tmpAranacak = Split(aranacak, ",")
tmpDegistir = Split(degistir, ",")

For z = 0 To UBound(tmpAranacak)

    bul = InStr(1, param1, tmpAranacak(z), vbTextCompare)
    If bul > 0 Then param1 = Replace(param1, tmpAranacak(z), tmpDegistir(z))
   
Next z

fncYazimDuzeni = param1
'bak = InStr(1, param, degisec, vbTextCompare)
'If bak > 0 Then

'fncYazimDuzeni=left(param,bak) &left(
End Function


Sub grafik_atanan()
Call baglan

SQL = "SELECT DISTINCT " _
& "siparis.siparis_no, yurtfirma.adi,siparis.siparis_tur_id " _
& "FROM siparis " _
& "INNER JOIN siparisdurum ON siparisdurum.siparis_id = siparis.id " _
& "INNER JOIN yurtfirma ON yurtfirma.id = siparis.firma_id " _
& "WHERE siparis.siparis_no like '18%' AND siparis.son_atanan_kullanici_id =12 AND siparis.son_durum_id =3"
'Debug.Print SQL

rs.Open SQL, con, 3, 4
lstSipler.Clear
i = 0
Do Until rs.EOF = True
    i = i + 1
    lstSipler.AddItem rs(0) & "|" & rs(1) & "|" & rs(2)
    rs.MoveNext
Loop

lblListe.Caption = "Toplam: " & i

'ba�lant�y� kapatal�m
Call kapat
End Sub



Private Sub Command1_Click()

'Pdfyi ta��yal�m
pdfkaynak = "\\172.16.11.27\2019 Takvim\2 Onay Bekliyor\190100986\190100986.pdf"
pdfhedef = "\\172.16.11.27\2019 Takvim\3 PDF\2 Renk\190100986x.pdf"

FileCopy pdfkaynak, pdfhedef
MsgBox "Pdf ta��nd�"

Exit Sub
grafik_bitenler
soru = InputBox("sipari� no", "ba�l�k", "180200999")

sipNo = soru
'Sipari� dizide varsa
If InStr(Join(sipListe), sipNo) > 0 Then
    
    'Sipari�in bilgilerini bulal�m
    For y = 0 To lstSipler.ListCount - 1
    
        tmpVeri = Split(lstSipler.List(y), "|")
                
        If tmpVeri(0) = sipNo Then
            MsgBox tmpVeri(0) & "*" & tmpVeri(1) & "*" & tmpVeri(2) & "*" & tmpVeri(3) & "*" & tmpVeri(4)
        End If
    Next y
End If


End Sub

Private Sub Command2_Click()
grafik_atanan

End Sub

Private Sub Form_Initialize()
'komut sat�r�n� okuyal�m
If Command$ <> "" Then
    komuts = Split(Trim(Command$), " ")
    For Each params In komuts
        If Left(params, 1) = "-" Then 'Parametre do�ru yaz�lm��sa i�lenir.
            tmpbol = Split(Replace(params, "-", ""), "=")
            dict.Add tmpbol(0), tmpbol(1)
        End If
    Next

End If

'Versiyon yazal�m
vers = App.Major & "." & App.Minor & "." & App.Revision
Me.Caption = " GEZGIN: v" & vers

'AYAR.ini dosyas� yolunu belirleyelim
fNo = FreeFile
yol = App.path & "\ayar.ini"

'Test ortam� �al���yorsa yol de�i�tirelim
If dict.Item("test") = "test" Then
    yol = App.path & "\ayar_test.ini"
    Me.Caption = Me.Caption & "-TEST"
End If

'Ayarlar� okuyup de�i�kenlere atal�m
Open yol For Input As #fNo
Do Until EOF(fNo)

    Line Input #fNo, satir
  If satir <> "" Then
    satirAyar = Split(satir, "=")
    dict.Add satirAyar(0), satirAyar(1)
  End If
  
i = i + 1
Loop
Close #fNo

'Loglar� okuyal�m
Call log_oku


'Varsay�lan de�erler
cboZaman.ListIndex = 0
dak = 0
san = 0
End Sub

Private Sub lblGrafiker_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #iFileNo
End Sub

Private Sub lblZaman_Click()

Select Case Timer1.Enabled
Case True
    Timer1.Enabled = False
    lblZaman.Caption = "00:00:00"
Case Else
    Timer1.Enabled = True
End Select


End Sub

Private Sub Timer1_Timer()
san = san + 1

'saniye 60 oldu�unda s�f�rla
If san = 60 Then
    san = 0
    dak = dak + 1
End If

lblZaman.Caption = Time


'E�ER ZAMAN 17:00:00 �SE OTOMASYONU DURDUR
'If lblZaman.Caption = "21:50:00" Then lblZaman_Click

'Sayac zaman� gelince i�lem yap
If dak = Val(cboZaman.Text) Then
    
    
    dak = 0
    
    cmdCalistir_Click
End If
End Sub

Private Sub Timer2_Timer()

say = say + 1
If say = 1 Then Timer2.Enabled = False
End Sub
