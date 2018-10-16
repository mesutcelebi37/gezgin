VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "AYARLAR"
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7470
   LinkTopic       =   "Form2"
   ScaleHeight     =   6525
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   555
      Left            =   1035
      TabIndex        =   2
      Top             =   4140
      Width           =   1815
   End
   Begin MSComctlLib.ListView lw2 
      Height          =   2715
      Left            =   450
      TabIndex        =   1
      Top             =   720
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   4789
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "1. Sutun"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "2 .Sutun"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "3.Sutun"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "bla bla bla"
      Height          =   420
      Left            =   1395
      TabIndex        =   0
      Top             =   45
      Width           =   1050
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      Height          =   330
      Left            =   4635
      TabIndex        =   4
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "1"
      Height          =   330
      Left            =   3015
      TabIndex        =   3
      Top             =   315
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itm As MSComctlLib.ListItem
Dim fnd As MSComctlLib.ListItem
Private Sub Command1_Click()
For i = 1 To 5
    lw2.ListItems.Add i, , i
    lw2.ListItems(i).ListSubItems.Add 1, , i & "inci deger"
    lw2.ListItems(i).ListSubItems.Add 2, , i & "inci deger"
    
Next

End Sub


Private Sub Command2_KeyPress(KeyAscii As Integer)
Call lw2_ItemCheck(itm)

End Sub

Private Sub Label1_Click()
Set fnd = lw2.FindItem(Label1.Caption, lwSubItem)

'Ürün bulunmuþsa
If Not fnd Is Nothing Then

    'Ürün önceden seçilmiþse
    If fnd.Checked = True Then
        Me.Caption = "'" & fnd & "' Zaten iþaretlenmiþ"

    Else
        
            
        fnd.EnsureVisible
        fnd.Checked = True
        fnd.Selected = True
        Me.Caption = "'" & Label1.Caption & "' iþaretlendi"
        Label2.Caption = Val(Label2.Caption) + 1
        
    
    End If
End If
End Sub

Private Sub lw2_ItemCheck(ByVal Item As MSComctlLib.ListItem)

If Item.Checked = True Then
    Label2.Caption = Val(Label2.Caption) + 1
Else
   Label2.Caption = Val(Label2.Caption) - 1
End If

End Sub



