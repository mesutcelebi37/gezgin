VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   225
      TabIndex        =   1
      Top             =   225
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Command1"
      Height          =   1050
      Left            =   2115
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Const SETTINGS_PROGID = "biopdf.PDFSettings"
Const UTIL_PROGID = "biopdf.PDFUtil"
Dim SQL, tmpVeri, tmpSipNo, tmpMiktar As String
Dim i As Integer

Private Function PrinterIndex(ByVal printerName As String) As Integer
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If LCase(Printers(i).DeviceName) Like LCase(printerName) Then
            PrinterIndex = i
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function

Private Sub cmdPrint_Click()
    Dim prtidx As Integer
    Dim sPrinterName As String
    Dim settings As Object
    Dim util As Object
    
    Set util = CreateObject(UTIL_PROGID)
    sPrinterName = util.defaultprintername
    
    Rem -- Configure the PDF print job
    Set settings = CreateObject(SETTINGS_PROGID)
    settings.printerName = sPrinterName
    settings.SetValue "Output", "<desktop>\myfile.pdf"
    settings.SetValue "ConfirmOverwrite", "no"
    settings.SetValue "ShowSaveAS", "never"
    settings.SetValue "ShowSettings", "never"
    settings.SetValue "ShowPDF", "yes"
    settings.SetValue "RememberLastFileName", "no"
    settings.SetValue "RememberLastFolderName", "no"
    settings.WriteSettings True
    
    Rem -- Find the index of the printer
    prtidx = PrinterIndex(sPrinterName)
    If prtidx < 0 Then Err.Raise 1000, , "No printer was found by the name of '" & sPrinterName & "'."
        
    Rem -- Set the current printer
    Set Printer = Printers(prtidx)
        
    Rem -- Print something
   ' If optOrientation(0).Value Then
        Printer.Orientation = PrinterObjectConstants.vbPRORPortrait
   ' Else
   '     Printer.Orientation = PrinterObjectConstants.vbPRORLandscape
   ' End If
    
    Rem -- Set paper size
    Rem -- http://msdn.microsoft.com/en-us/library/microsoft.visualbasic.powerpacks.printing.compatibility.vb6.printer.papersize.aspx
    Rem -- Note: Custom paper size is not supported by VB6 after Windows 98.
    'Printer.PaperSize = vbPRPSB5
    
    Call yeni_sipler
    
    'Yeni sipariþ yoksa çýkalým
    If List1.ListCount < 0 Then Exit Sub
    
    
    'mavi yazalým
    Printer.ForeColor = vbBlue
    
    'Yeni sipariþler kadar gezintiye çýkalým
    For i = 0 To List1.ListCount - 1
        'TMPYOL DOSYASINI BULALIM
        
        'Dosya sisteminde yeni sipariþ için klasör var mý bakalým
        If Dir() <> "" Then

        End If
        tmpVeri = Split(List1.List(i), "|")
        tmpSipNo = tmpVeri(0)
        tmpMiktar = tmpVeri(1)
        
        Printer.FontSize = 10
        Printer.Print tmpSipNo
        Printer.Print tmpMiktar
                        

                
                
    Next i
    

    
    Printer.FontSize = 10
        
    'Printer.Print "The time is " & Now
    Printer.EndDoc
    
    Rem -- Wait for runonce settings file to disappear
    Dim runonce As String
    runonce = settings.GetSettingsFilePath(True)
    While Dir(runonce, vbNormal) <> ""
        Sleep 100
    Wend
    
    'MsgBox "myfile.pdf was saved on your desktop", vbInformation, "PDF Created"
End Sub


Sub yeni_sipler()
Call baglan


SQL = "SELECT DISTINCT " _
& "siparis.siparis_no AS siparis_no,siparis.toplam_adet,siparis.onay_iste " _
& "FROM " _
& "siparis " _
& "Where " _
& "siparis.son_durum_id IN (3) AND " _
& "siparis.siparis_no LIKE '180101589'"

rs.Open SQL, con, 3, 4

i = 0
Do Until rs.EOF = True
    i = i + 1
    List1.AddItem rs(0) & "|" & rs(1) & "|" & rs(2)
    rs.MoveNext
Loop


rs.Close
con.Close

End Sub

