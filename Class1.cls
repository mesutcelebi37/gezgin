VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "crud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub baglan()
    'On Error GoTo errHandler:
    Set con1 = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    con1.Open = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\term.mdb;Persist Security Info=False"
    
    
End Sub
