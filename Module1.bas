Attribute VB_Name = "Module1"
Global dict As New Scripting.Dictionary 'Ayarlarý kontrol etmek için
Global vers As String 'Program versiyonunu takip etmek için

Global con As New ADODB.Connection
Global rs As New ADODB.Recordset

Dim iFileNo As Integer

'Dosyaya log ekleme fonksiyonu
Function log_ekle(tip As String, kullanici As String, nott As String)

    iFileNo = FreeFile
    If tip = "" Then tip = "Bilgi"
    If usr = "" Then usr = "Genel"
    
    yol = dict.Item("logKlasor") & "\" & Format(Date, "YY-mm") & "-Log.csv"
    If Dir(yol) <> "" Then 'Dosya varsa ekliyor. Yoksa Açýyor.
        Open yol For Append As #iFileNo
    Else
        Open yol For Output As #iFileNo
    End If
    

    'Log Formatý Versiyon | Tip("Bilgi,Uyarý,Hata") | Kullanýcý | Tarih | Kullanýcý | Not
    'Gelen Notu yazalým
    notyaz = "v" & vers & ";" & tip & ";" & Now & ";" & kullanici & ";" & nott
    Print #iFileNo, notyaz
    
    'frmAna.lstDurum.AddItem notyaz
    
    Close #iFileNo
    
'Son logu okuyalým
Call log_oku
End Function

Function log_oku()
    iFileNo = FreeFile
    'Listeyi temizleyelim
    frmAna.lstDurum.Clear
    yol = dict.Item("logKlasor") & "\" & Format(Date, "YY-mm") & "-Log.csv"
    If Dir(yol) <> "" Then
        Open yol For Input As #iFileNo
        Do While Not EOF(iFileNo)
            Line Input #iFileNo, tmpStr
            frmAna.lstDurum.AddItem tmpStr
        Loop
    
    End If
    
    Close #iFileNo
    
'Listbox ta son deðere konumlanalým
If frmAna.lstDurum.ListCount > 0 Then frmAna.lstDurum.ListIndex = frmAna.lstDurum.ListCount - 1

End Function

Function baglan()
con.Open ("DRIVER={MySQL ODBC 5.2 ANSI Driver};" _
& "SERVER=45.76.91.110;" _
& "DATABASE=faziletc_siparis;" _
& "USER=faziletc_siparis;" _
& "PASSWORD=Hizmet4445960;" _
& "OPTION=3;")

End Function

Function kapat()
If rs.State = 1 Then rs.Close
If con.State = 1 Then con.Close
End Function
