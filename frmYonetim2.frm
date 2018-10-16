VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmYonetim2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TERM�NAL ADM�N"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2520
   ForeColor       =   &H80000018&
   Icon            =   "frmYonetim2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   2520
   Begin VB.ComboBox cboPrg 
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
      ItemData        =   "frmYonetim2.frx":1C64C
      Left            =   45
      List            =   "frmYonetim2.frx":1C656
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1755
      Width           =   1095
   End
   Begin VB.ComboBox cboVers 
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
      ItemData        =   "frmYonetim2.frx":1C66B
      Left            =   1305
      List            =   "frmYonetim2.frx":1C66D
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "Vers"
      Top             =   1755
      Width           =   1095
   End
   Begin VB.TextBox txtAciklama 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmYonetim2.frx":1C66F
      Top             =   405
      Width           =   2355
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "VERS. B�LG�S� EKLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2070
      Width           =   2355
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   540
      TabIndex        =   3
      Top             =   90
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      _Version        =   393216
      Format          =   98369537
      CurrentDate     =   42889
   End
End
Attribute VB_Name = "frmYonetim2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Private Declare Function Beep Lib "kernel32" (ByVal soundFrequency As Long, ByVal soundDuration As Long) As Long
Private Declare Function Beep Lib "kernel32" (ByVal soundFrequency As Long, ByVal soundDuration As Long) As Long


Dim itm As MSComctlLib.ListItem

Dim lstHar(0 To 20) As String 'Hareket kodlar� verilerini tutmak i�in

Dim i, fNo, secilen, sayac As Integer
Dim sipHazir, sipHazirlanacak, sipTumu, lw_sira As Integer
Dim sipno, sipMasterDetayNo As String
Dim bir_kere, bulundu As Boolean

'VERS�YONLAR
'v3.0.179:  1. Faz g�ncellemeleri bitti
'v3.0.180:  Versiyon sistemine ge�ildi
'v3.0.376:  Hareket Kodlar� TERMINALDB'den al�n�yor



Private Sub chkIsaret_Click()
If chkIsaret.Value = 1 Then
    chkIsaret.Caption = "+"
Else
    chkIsaret.Caption = "-"
End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkUyum_Click()
If chkUyum.Value = 1 Then
    frmUyum.Visible = True
Else
    frmUyum.Visible = False
End If
End Sub

Private Sub cmd_musteri_Click()
DoEvents
cmd_musteri.Visible = False

lw1.ListItems.Clear
'BA�LANTI AYARLANIYOR
con1_baglan

EK_SQL = ""
parca_must = Split(cboMust.Text, "|")

If cboMust.ListIndex <> 0 Then EK_SQL = " AND musteri_kod='" & parca_must(1) & "'"

parca_harKod = Split(cboHarkod.Text, "|")
If cboHarkod.ListIndex <> 0 Then EK_SQL = " AND harkod='" & parca_harKod(0) & "'"

tar1 = Format(DTPicker3.Value, "m-d-YYYY")
tar2 = Format(DTPicker4.Value, "m-d-YYYY")

SQL = "SELECT * FROM sipler WHERE (sipTarih BETWEEN #" & tar1 & "# AND #" & tar2 & "#) " & EK_SQL
Text1.Text = SQL
rs.Open SQL, con1, adOpenStatic, adLockReadOnly
For i = 1 To rs.RecordCount

        Set lwsatir = lw1.ListItems.Add(, , rs![ID])
       
        
        lwsatir.ListSubItems.Add 1, , Trim(rs![sipno])
        lwsatir.ListSubItems.Add 2, , Trim(rs![musteri_ad])
        lwsatir.ListSubItems.Add 3, , Trim(rs![tutar])
       
        lwsatir.ListSubItems.Add 4, , Trim(rs![sipTarih])
        
        If IsNull(rs![tesTarih]) = True Then
            lwsatir.ListSubItems.Add 5, , ""
        Else
            lwsatir.ListSubItems.Add 5, , Trim(rs![tesTarih])
        End If

        lwsatir.ListSubItems.Add 6, , Trim(rs![harkod])
   
        If IsNull(rs![Term]) = True Then
            lwsatir.ListSubItems.Add 7, , ""
        Else
            lwsatir.ListSubItems.Add 7, , Trim(rs![Term])
        End If
        
    rs.MoveNext
Next i

'BA�LANTI KAPATILIYOR.
con1_kapat
cmd_musteri.Visible = True


lw1.SetFocus
Call urunleri_goster




End Sub


Private Sub cmdAra_Click()

lw1.ListItems.Clear
Call acik_sipler

log_ekle (Now & "= '" & cboMust.Text & "' B�lgesi arand�")
txtBarkod.SetFocus
End Sub


Private Sub cmd_uyumdan_Click()
DoEvents

'On Error GoTo errhandler

cmd_uyumdan.Visible = False


'UYUMA BA�LANALIM
con_baglan

'Se�ilen sipari� detaylar� listeleniyor.
SQLX = "SELECT " _
& "cari_ad, mal_tutar, sip_no, hareket_kod, sip_tarih, teslim_tarih " _
& "FROM PUB.siparis_detay " _
& "" _
& "WHERE sip_tarih ='" & Format(DTPicker1.Value, "YYYY-mm-dd") & "' AND firma_kod='" & firma & "' AND siparis_durumu=1 AND siparis_master.sip_no ='KT-1704165'"

SQL = " SELECT " _
& "siparis_detay.stok_kod,siparis_detay.dmiktar,siparis_detay.sip_no,siparis_detay.firma_kod,siparis_detay.sip_detayno,siparis_detay.sip_masterno,siparis_detay.stok_ad,siparis_detay.dbirim,siparis_detay.siparis_durum," _
& "stok_barkod.Bar_kod,stok_barkod.sira_no," _
& "depo_stok.raf_kod," _
& "siparis_master.sip_no, siparis_master.firma_kod,siparis_master.cari_kod,siparis_master.cari_ad,siparis_master.mal_tutar,siparis_master.hareket_kod,siparis_master.sip_tarih, siparis_master.teslim_tarih " _
& "FROM PUB.siparis_detay " _
& "LEFT OUTER JOIN PUB.siparis_master ON siparis_detay.firma_kod=siparis_master.firma_kod AND siparis_detay.sip_masterno=siparis_master.sip_masterno " _
& "LEFT OUTER JOIN PUB.stok_barkod ON siparis_detay.firma_kod=stok_barkod.firma_kod AND siparis_detay.stok_kod=stok_barkod.stok_kod AND siparis_detay.dbirim=stok_barkod.dbirim " _
& "LEFT OUTER JOIN PUB.depo_stok ON siparis_detay.firma_kod=depo_stok.firma_kod AND siparis_detay.stok_kod=depo_stok.stok_kod AND siparis_detay.depo_kod=depo_stok.depo_kod " _
& "WHERE (siparis_master.sip_tarih BETWEEN '" & Format(DTPicker1.Value, "YYYY-mm-dd") & "' AND '" & Format(DTPicker2.Value, "YYYY-mm-dd") & "') AND siparis_master.firma_kod='" & firma & "' AND siparis_master.siparis_durumu=1 AND stok_barkod.sira_no=1"

SQL1 = "SELECT * FROM urunler"
SQL2 = "SELECT * FROM sipler"
urun_say = 0
siparis_say = 0

'UYUMSOFT ���N RECORDSET
rs.Open SQL, con, adOpenForwardOnly, adLockReadOnly

'KAYIT BULAMAZSA �IKIYOR.
If rs![sip_no] = "" Then
    con_kapat
    Exit Sub
End If

'TERMINAL DBYE BA�LANALIM
con1_baglan


'�R�NLER TABLOSU ���N RECORDSET
rs1.Open SQL1, con1, adOpenStatic, adLockOptimistic

'S�PAR��LER TABLOSU ���N RECORDSET
rs2.Open SQL2, con1, adOpenStatic, adLockOptimistic



'UYUMDAN BULUNAN VER�LERDE GEZERKEN �R�NLER TABLOSUNA VE S�PAR�� TABLOSUNA VER�LER� EKL�YOR.
Do Until rs.EOF = True

    'S�PAR�� DE����RSE S�PAR�� TABLOSUNA SATIR VER�S�N�N EKLENMES� GEREK�YOR.
    If sip_no_temp <> rs![sip_no] Then
    
        'Burada sipari�in al�n�p al�nmad���n� kontrol edip gerekirse pas ge�mek laz�m
        SQL3 = "SELECT sipno FROM sipler WHERE sipno='" & rs![sip_no] & "'"
        rs3.Open SQL3, con1, adOpenStatic, adLockReadOnly
       
       
        If rs3.RecordCount > 0 Then
        rs3.Close
    GoTo pas_gec:
        
        End If
        'HER SATIRDA �ALI�ACAK RS3 KAPATILIYOR.
        rs3.Close
    
        siparis_say = siparis_say + 1
        rs2.AddNew
        rs2![musteri_kod] = rs![cari_kod]
        rs2![musteri_ad] = rs![cari_ad]
        rs2![tutar] = rs![mal_tutar]
        rs2![sipno] = rs![sip_no]
        rs2![harkod] = rs![hareket_kod]
        rs2![sipTarih] = rs![sip_tarih]
        rs2![tesTarih] = rs![teslim_tarih]
        rs2.Update
        
    End If
  
    rs1.AddNew
    rs1![stok_kod] = rs![stok_kod]
    rs1![stok_ad] = rs![stok_ad]
    
    rs1![Bar_kod] = rs![Bar_kod]
    rs1![dmiktar] = rs![dmiktar]
    rs1![dbirim] = rs![dbirim]

    rs1![raf_kod] = rs![raf_kod]
    rs1![sip_masterno] = rs![sip_masterno]
    rs1![sip_detayno] = rs![sip_detayno]
    rs1![sip_no] = rs![sip_no]
    rs1![durum] = 0
    rs1.Update


sip_no_temp = rs![sip_no]
urun_say = urun_say + 1

pas_gec:
rs.MoveNext
Loop

cmd_uyumdan.Visible = True

'BA�LANTILARI KAPAT
con_kapat
con1_kapat


fnc_log ("Terminal DD ye al�nan, SIPARIS: " & siparis_say & ", �R�N: " & urun_say)
MsgBox siparis_say & " tane sipari� Uyumdan Al�nd�."



Exit Sub
errhandler:
MsgBox Err.Description

'BA�LANTILARI KAPATALIM
con_kapat
con1_kapat
End

End Sub

Private Sub cmdIrs_Click()
'On Error GoTo errHandler:
soru = MsgBox("Onaylay�n�z", vbYesNo + vbDefaultButton2, "�RSAL�YE ONAYI")

'��LEM� ONAYLAMAZSA �IKIYOR.
If soru <> 6 Then Exit Sub

'BA�LANALIM
con1_baglan

SQL = "SELECT * FROM sipler " _
& "INNER JOIN urunler ON sipler.sipno=urunler.sip_no " _
& "WHERE sipler.term='" & terminal & "' " _
& "AND urunler.durum=1"

'Debug.Print SQL
rs.Open SQL, con1, adOpenStatic, adLockOptimistic

'TERM�NAL�N HAZIRLADI�I �R�N YOKSA
If rs.RecordCount = 0 Then
    MsgBox "�rsaliye kesecek �r�n bulunamad�."
    'BA�LANTI KAPATILIYOR.
    con1_kapat
    Exit Sub
End If

'data dosyas� olu�turuluyor.*******************************
fNo = FreeFile
yol = App.Path & "\yedek\data1_" & terminal & "_" & Format(Now, "dd-m-YYYY H-mm") & ".txt"
Open yol For Output As #fNo

For i = 1 To rs.RecordCount
    
   satir = """" & rs![stok_kod] & """;" & rs![dmiktar] & ";" & rs![sip_masterno] & ";" & rs![sip_detayno] & ";" & Format(Date, "dd/mm/YYYY") & ";""" & rs![sip_no] & """;""" & rs![harkod] & """;""H"";;""" & terminal & "-v:" & vers & """"
    'Sat�r ekleniyor.
   Print #fNo, satir

'DURUMU G�NCELLE
rs![durum] = 2
rs.Update
rs.MoveNext
Next i


MsgBox "��lem Tamamland�. Hemen UYUM'dan irsaliye kesiniz!!!", vbCritical


'Dosya kapat�l�yor.
Close #fNo


'BA�LANTI KAPATILIYOR.
con1_kapat

'Log ekleniyor.
'fnc_log ("�RSAL�YE DOSYASI HAZIRLANDI")


'Hata durumunda
Exit Sub
    
errhandler:
    MsgBox Err.Description, vbCritical, "HATA"

End Sub

Private Sub cmdKaydet_Click()
'LW BO� �SE �IKIYOR.
If lw1.ListItems.Count = 0 Then Exit Sub


sip_term = lw1.SelectedItem.ListSubItems(5).Text
sip_id = lw1.SelectedItem.Text
sip_no = lw1.SelectedItem.ListSubItems(1).Text
musteri = lw1.SelectedItem.ListSubItems(2).Text
siptar = lw1.SelectedItem.ListSubItems(3).Text

If sip_term <> "" Then

    If sip_term <> terminal Then
        MsgBox "HATA:Ba�ka Terminal Zaten Haz�rl�yor.", vbCritical
        Exit Sub
    Else
        MsgBox "Yar�m kalan sipari�e devam"
    End If
    
End If

'BAGLAN
con1_baglan

'S�PAR��� TERM�NALE AL
SQL = "UPDATE sipler SET term='" & terminal & "' WHERE ID=" & sip_id & ""
con1.Execute (SQL), etkilenen

'BA�LANTIYI KAPAT
con1_kapat

If etkilenen = 1 Then
    'LOG EKLE
    fnc_log (sip_no & " nolu sipari� haz�rlanacak")
    
    Form2.Show 1
End If



Exit Sub
'HATA OLU�TU�UNDA
errhandler:
MsgBox Err.Description, vbCritical, "HATA"

End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Sub urunleri_goster()
DoEvents
'Sipari� se�ilmemi�se ��k�yor.
If lw1.ListItems.Count = 0 Then Exit Sub
'L�STV�EW SIFIRLANIR.
lw2.ListItems.Clear


'BA�LANALIM
con1_baglan

SQL = "SELECT * FROM urunler WHERE sip_no='" & lw1.SelectedItem.ListSubItems(1).Text & "' ORDER BY raf_kod"
rs.Open SQL, con1, adOpenStatic, adLockReadOnly

For i = 1 To rs.RecordCount

    Set lwsatir2 = lw2.ListItems.Add(, , rs![ID])
'If Trim(rs![durum]) = 2 Then lwsatir2.ListSubItems(i).FontBold = True
    lwsatir2.ListSubItems.Add 1, , Trim(rs![stok_kod])
    lwsatir2.ListSubItems.Add 2, , Trim(rs![stok_ad])
    lwsatir2.ListSubItems.Add 3, , Trim(rs![dmiktar])
    lwsatir2.ListSubItems.Add 4, , Trim(rs![dbirim])
    If IsNull(rs![raf_kod]) = True Then
        lwsatir2.ListSubItems.Add 5, , ""
    Else
        lwsatir2.ListSubItems.Add 5, , Trim(rs![raf_kod])
    End If
    
     
    lwsatir2.ListSubItems.Add 6, , mod_fnc_durum(Trim(rs![durum]))
    If rs![durum] = 1 Then lwsatir2.ListSubItems(6).ForeColor = &HFF&
    rs.MoveNext
Next i

'lw2.Refresh
'Ba�lant�y� kapatal�m
con1_kapat



End Sub

Private Sub cmdRota_Click()
If lw2.ListItems.Count = 0 Then Exit Sub
sayac = 0

For i = 1 To lw2.ListItems.Count
    sokak = Trim(Left(lw2.ListItems(i).ListSubItems(5).Text, 1))
       

    'SOKAK L�STEDE YOKSA
    If InStr(1, sokaklar, sokak, vbTextCompare) = 0 And sokak <> "" Then
        sokaklar = sokaklar & sokak
        sayac = sayac + 1
        rapor = rapor & sayac & ") '" & sokak & "' Soka��na gidiniz" & vbCrLf

    End If
    
Next i

'RAPORU G�STER
MsgBox rapor


End Sub

Private Sub Command1_Click()
'YAPILACAKLAR
'1 S�PAR��LER UYUMDAN �EK�LECEK VE TERM�NAL DBYE Y�KLENECEKT�R.
'S�PAR�� DAHA �NCE Y�KLENM��SE Y�KLEMEYECEKT�R.
'TAR�H BAZLI VE HAREKET KODU BAZLI S�PAR��LER Y�KLENECEKT�R.
'SQL SORGUSU RAPORDAN ALINACAKTIR.
'S�PAR�� VE �R�NLER�N Y�KLENMES� �EKL�NDE �K� PAR�ADA YAPILACAKTIR.

'T�M�N� SE� VE  BENZER� F�LTRELEMELER OLMALIDIR.
'TERM�NALDEK� S�PAR��LER� S�LECEK B�R YAPI OLMALIDIR.
parcala = Split(cboHarkod.Text, "|")
MsgBox parcala(0)
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
'BA�LANALIM
con1_baglan

SQL = "SELECT * FROM log ORDER BY tarih desc"
rs1.Open SQL, con1, adOpenStatic, adLockReadOnly



If rs1.RecordCount = -1 Then Exit Sub

For i = 1 To rs1.RecordCount
    rapor = rapor & rs1![tarih] & ":" & rs1![aciklama] & vbCrLf
    
    If i = 15 Then Exit For
    rs1.MoveNext
Next i

MsgBox rapor

'BA�LANTIYI KAPATALIM
con1_kapat

End Sub

Private Sub Command4_Click()
MsgBox terminal

End Sub

Private Sub cmdLog_Click()

If cboPrg.ListIndex = -1 Then Exit Sub

'BAGLAN
con1_baglan

'VERS�YON A�IKLAMASI EKLE
zaman = Format(DTPicker1.Value, "m-d-YYYY") & " " & Format(Time, "HH:mm:ss")
SQL = "INSERT INTO versiyon (prog, tarih, vers, aciklama) VALUES ('" & cboPrg.Text & "','" & zaman & "','" & cboVers.Text & "','" & txtAciklama.Text & "')"
con1.Execute SQL

'BA�LANTIYI KAPAT
con1_kapat

MsgBox "Versiyon a��klamas� eklendi"




Exit Sub
'HATA OLU�TU�UNDA
errhandler:
cmdKaydet.Visible = True
MsgBox Err.Description, vbCritical, "HATA"

End Sub

Private Sub hareket_kodlari()
'BA�LANTIYI A�ALIM
con1_baglan

'HAraket kodlar�n� doldural�m
SQL = "SELECT DISTINCT harkod, ayarlar.degeri FROM sipler LEFT JOIN ayarlar ON ayarlar.kodu=sipler.harkod"

rs1.Open SQL, con1, adOpenStatic, adLockReadOnly


For i = 1 To rs1.RecordCount
'    If rs1![adi] = "harkod" Then
           cboHarkod.AddItem rs1![harkod] & "|" & rs1![degeri]
    'End If
    
rs1.MoveNext
Next i
rs1.Close

'BA�LANTIYI KAPATALIM
con1_kapat
End Sub
Private Sub acik_musteriler()
'lblBilgi.Caption = ""

'On Error GoTo errHandler:
'BA�LANALIM
con1_baglan

SQL = "SELECT DISTINCT musteri_kod, musteri_ad FROM sipler"
rs.Open SQL, con1, adOpenStatic, adLockReadOnly
sayac = 0
i = 1

'KAYIT BULUNAMAZ �SE
If rs.RecordCount = 0 Then
    MsgBox "Sipari� Bulunamad�"
    'BA�LANTIYI KAPAT
    con1_kapat
    Exit Sub
End If

For i = 1 To rs.RecordCount
    

   cboMust.AddItem rs![musteri_ad] & "|" & rs![musteri_kod]
  ' cboMust.ItemData(cboMust.NewIndex) = rs![musteri_ad] & "|" & rs![musteri_kod]

  
    rs.MoveNext
Next

'RS ve ba�lant� kapat�l�yor.
con1_kapat

Exit Sub
'HATA DURUMUNDA
errhandler:
MsgBox "HATA:" & Err.Description
End Sub

Private Sub acik_sipler()
lblBilgi.Caption = ""

'On Error GoTo errHandler:

'BA�LANTI AYARLANIYOR.
con1_baglan

'B�lge se�ilmi�se
SQL = ""
If cboMust.ListIndex <> -1 Then
    veri = Split(cboMust.Text, "|")
    EK_SQL = " AND siparis_master.cari_kod='" & veri(1) & "' "
End If

'Se�ilen sipari� detaylar� listeleniyor.
SQL = "SELECT " _
& "siparis_detay.firma_kod, siparis_detay.stok_kod, siparis_detay.stok_ad, siparis_detay.dmiktar, siparis_detay. sip_masterno, siparis_detay.sip_detayno, " _
& "siparis_master.sip_no, siparis_master.aciklama1, siparis_master.siparis_durumu, siparis_master.teslim_tarih " _
& "FROM PUB.siparis_detay INNER JOIN PUB.siparis_master ON siparis_detay.firma_kod = siparis_master.firma_kod AND siparis_detay.sip_masterno = siparis_master.sip_masterno " _
& "WHERE siparis_detay.firma_kod='CAM2017' AND siparis_master.siparis_durumu = 1 AND siparis_master.sip_no like 'KT-1704182'"

SQL2 = "SELECT siparis_durumu FROM PUB.siparis_master WHERE firma_kod='CAM2017' AND sip_no like 'KT-1704182'"


Set rs = New ADODB.Recordset
rs.Open SQL2, con, adOpenStatic, adLockReadOnly

MsgBox rs.RecordCount
Exit Sub

sayac = 0

i = 1

'KAYIT BULUNAMAZ �SE
If rs.RecordCount = 0 Then
    MsgBox "Kay�t Bulunamad�1"

    'BA�LANTIYI KAPAT
    con1_kapat
    Exit Sub
End If



For i = 1 To rs.RecordCount
    
    
    'YEN� S�PAR�� NUMARASINDA SAYA� ARTIYOR.
    If rs![sip_no] <> sipno Then
        sayac = sayac + 1
    End If
    sipno = rs![sip_no]
    
        Set lwsatir = lw1.ListItems.Add(, , 0)
        lwsatir.ListSubItems.Add 1, , Trim(rs![sip_no])
        lwsatir.ListSubItems.Add 2, , Trim(rs![dmiktar])
        lwsatir.ListSubItems.Add 3, , Trim(rs![aciklama1])
        lwsatir.ListSubItems.Add 4, , Trim(rs![teslim_tarih])
        lwsatir.ListSubItems.Add 5, , Trim(rs![stok_kod])
                
        '5 SONRASI KAYIT ���N GEREKENLER
        lwsatir.ListSubItems.Add 6, , Trim(rs![stok_kod])
        lwsatir.ListSubItems.Add 7, , Trim(rs![dmiktar])
        lwsatir.ListSubItems.Add 8, , Trim(rs![sip_masterno])
        lwsatir.ListSubItems.Add 9, , Trim(rs![sip_detayno])
   
    rs.MoveNext
Next
Me.Caption = rs.RecordCount & " Detay bulundu."

'RS ve ba�lant� kapat�l�yor.
con1_kapat


'LW SIRALANIYOR.
lw1.SortKey = 1
lw1.Sorted = True

'HAZIRLANAN S�PAR��LER ��ARETLEN�YOR.
yol = App.Path & "\data1_" & dict.Item("Term") & ".txt"

fNo = FreeFile
i = 1
Open yol For Input As #fNo
Do Until EOF(fNo)

    Line Input #fNo, satir_veri
    satir = Split(Replace(satir_veri, """", ""), ";")
        
        'Her sat�r i�in lwde d�n�yor.
        For Y = 1 To lw1.ListItems.Count
            If lw1.ListItems(Y).ListSubItems(1).Text = Replace(satir(5), """", "") Then
                
                lw1.ListItems(Y).Text = 2
            lw1.ListItems(Y).Checked = True
            End If
    
        Next Y
        
      
i = i + 1
Loop
Close #fNo

'Gerekli Bilgiler dolduruluyor
sipTumu = sayac
sipHazir = 0
lblBilgi.Caption = sipHazir & "/" & sipTumu



'HATALAR
Exit Sub
errhandler:
MsgBox "HATA:" & Err.Description
End Sub
Private Sub urun_bul()

Set fnd = lw1.FindItem(txtBarkod.Text, lvwText)

'�r�n bulunmu�sa
If Not fnd Is Nothing Then

    '�r�n �nceden se�ilmi�se
    If fnd.Checked = True Then
        Me.Caption = "'" & txtBarkod.Text & "' zaten i�aretlenmi�"
    Else
        fnd.EnsureVisible
        fnd.Checked = True
        fnd.ListSubItems(1).Bold = True
        fnd.Selected = True
        'lw1.SetFocus
        Me.Caption = "'" & txtBarkod.Text & "' i�aretlendi"
        sipHazir = sipHazir + 1
        lblBilgi.Caption = sipHazir & "/" & sipTumu

    End If
    

Else
    Me.Caption = "'" & txtBarkod.Text & "' Bulunamad�"

    
End If

txtBarkod.Text = ""
End Sub
Private Sub siparis_bul_isaretle()

bir_kere = False
bulundu = False

For i = 1 To lw1.ListItems.Count
        
    'Sipari� listede bulunduysa
    If txtBarkod.Text = lw1.ListItems(i).ListSubItems(1).Text Then
        
        bulundu = True
        'Bir kereye mahsus kontroller yap�l�yor.
        If bir_kere = False Then
            bir_kere = True
                
            'Daha �nceden i�aretlenmi� ise ��k�yor
            If lw1.ListItems(i).Text = 2 Then
                Me.Caption = "'" & txtBarkod.Text & "' �nceden haz�rlanm��"
                Beep 250, 250
                Beep 250, 250
                Exit Sub
            End If
            
            '�imdiki aramada listede i�aretlenmi� ise ��k�yor
            If lw1.ListItems(i).Checked = True And opArti.Value = True Then
                Me.Caption = "'" & txtBarkod.Text & "' zaten i�aretlenmi�"
                Beep 250, 250
                Beep 250, 250
                Exit Sub
            End If
            
            'BULDU�U ���N UYARI SES� �IKARIYOR.
            Beep 250, 250
        End If
        
        'Bulunan sipari� detaylar�nda geziliyor.
        'Ters i�aret i�lemi
        If opEksi.Value = True Then
            lw1.ListItems(i).Text = 0
            lw1.ListItems(i).Checked = False
            lw1.ListItems(i).ListSubItems(1).Bold = False
        
        Else
            '��aretleme i�lemi
            lw1.ListItems(i).Text = 1
            lw1.ListItems(i).Checked = True
            lw1.ListItems(i).Selected = True
            lw1.ListItems(i).EnsureVisible
            
            lw1.ListItems(i).ListSubItems(1).Bold = True
        End If
               
        'Sipari� miktar� i�in
        If Left(lw1.ListItems(i).ListSubItems(5).Text, 2) <> "60" Then miktar = miktar + Val(lw1.ListItems(i).ListSubItems(2).Text)

    End If
    
Next

'�R�N�N BULUNDU�U VE BULUNMADI�I DURUMLARDA B�R KERE YAPILACAKLAR
If bulundu = True Then
    'Ters i�aretleme durumu
    If opEksi.Value = True Then
        Me.Caption = "'" & txtBarkod.Text & "' " & miktar & " adet i�aret kald�r�ld�"
        If sipHazir > 0 Then sipHazir = sipHazir - 1
        opArti.Value = True
    Else
        '��aretleme durumu
        Me.Caption = "'" & txtBarkod.Text & "' " & miktar & " adet i�aretlendi"
        sipHazir = sipHazir + 1
        
    End If
    

    lblBilgi.Caption = sipHazir & "/" & sipTumu
Else
    Me.Caption = "'" & txtBarkod.Text & "' Bulunamad�"
End If

txtBarkod.Text = ""
txtBarkod.SetFocus
End Sub

Private Sub lw1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lw_sira = ColumnHeader.Index - 1
lw1.SortKey = lw_sira
lw1.Sorted = True
End Sub

Private Sub lw1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Checked = True And Item.Text = 0 Then
    Item.Text = 1
End If
End Sub

Private Sub opArti_Click()
    txtBarkod.SetFocus
End Sub

Private Sub opEksi_Click()
    txtBarkod.SetFocus
End Sub

Private Sub txtBarkod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Me.Caption = ""
    Call siparis_bul_isaretle
    End If

End Sub

Function log_ekle(mesaj As String)
    'AYAR DOSYASI OKUNUYOR
    fNo = FreeFile
    yol = App.Path & "\log\" & Format(Date, "YYYY-mm-dd") & ".txt"
    
    If Dir(yol) = "" Then
        Open yol For Output As #fNo
    Else
        Open yol For Append As #fNo
    End If
        'Mesaj ekleniyor
        Print #fNo, mesaj

    Close #fNo
End Function

Private Sub lw1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Call urunleri_goster
End Sub

Private Sub Form_Load()
cboVers.AddItem vers
cboVers.ListIndex = 0
cboPrg.ListIndex = frmYonetim.cboPrg.ListIndex
DTPicker1.Value = Format(Date, "dd/mm/YYYY")
End Sub
