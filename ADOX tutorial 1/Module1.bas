Attribute VB_Name = "Module1"
Public cn As Connection
Public cat As Catalog
Public i As Integer
Public rs As Recordset
Public selectedTbl As Table

Sub init()
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set cat = New Catalog
End Sub

