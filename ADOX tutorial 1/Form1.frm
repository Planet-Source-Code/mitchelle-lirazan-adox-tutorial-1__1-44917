VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADOX Sample 1"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   885
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   60
      Width           =   4275
   End
   Begin VB.ListBox List2 
      Height          =   2205
      Left            =   2520
      TabIndex        =   2
      Top             =   1620
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "browse mdb file.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   600
      TabIndex        =   1
      Top             =   1020
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   570
      TabIndex        =   0
      Top             =   1620
      Width           =   1545
   End
   Begin MSComDlg.CommonDialog cmdg 
      Left            =   3330
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.mdb"
      DialogTitle     =   "Select mdb file"
      FileName        =   "*.mdb"
   End
   Begin VB.Label Label2 
      Caption         =   "Fields:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2550
      TabIndex        =   4
      Top             =   1410
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Tables:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   1410
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
On Error GoTo errH
    cmdg.CancelError = False
    cmdg.ShowOpen
    
    'check if the connection is already open. Close if it is open
    If cn.State = 1 Then cn.Close
    
    cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & cmdg.FileName
    Set cat.ActiveConnection = cn
    
    'loop thru tables
    List1.Clear
    For i = 0 To cat.Tables.Count - 1
        If LCase(cat.Tables(i).Type) = "table" Then
            List1.AddItem cat.Tables(i).Name
        End If
    Next
errH:
    Dim pw As String
    'check if there is a database password
    If Err.Number = -2147217843 Then
        Form2.Show
    End If
End Sub




Private Sub Form_Load()
    Text1 = "This ADOX sample demonstrate how to open an mdb file " & vbCrLf & _
            "and displays the tables it contains in the Tables list, " & vbCrLf & _
            "and as you select a table in Tables list the fields " & vbCrLf & _
            "it contains will be listed in Fields list."
    init
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub List1_Click()
Dim tbl As New Table
    List2.Clear
    
    For i = 0 To cat.Tables.Count - 1
        If LCase(cat.Tables(i).Name) = LCase(List1.Text) Then Exit For
    Next
    Set selectedTbl = cat.Tables(i)
    Set tbl = cat.Tables(i)
    'list all fields of the selected table
    For i = 0 To tbl.Columns.Count - 1
        List2.AddItem tbl.Columns(i).Name
    Next
    
End Sub


