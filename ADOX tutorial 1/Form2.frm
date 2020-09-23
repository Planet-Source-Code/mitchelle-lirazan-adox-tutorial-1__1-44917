VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Input Password"
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   LinkTopic       =   "Form2"
   ScaleHeight     =   1515
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   405
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   405
      Left            =   1140
      TabIndex        =   2
      Top             =   840
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   330
      Width           =   2235
   End
   Begin VB.Shape Shape1 
      Height          =   1365
      Left            =   60
      Top             =   90
      Width           =   4245
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   390
      TabIndex        =   1
      Top             =   360
      Width           =   1305
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    On Error GoTo errH
    'check if the connection is already open. Close if it is open
    If cn.State = 1 Then cn.Close
    
    cn.Provider = "microsoft.jet.oledb.4.0;jet oledb:database password=" & Trim(Text1.Text)  'jet oledb:database password=12345
    cn.Open Form1.cmdg.FileName
    
    
    Set cat.ActiveConnection = cn
    'loop thru tables
    Form1.List1.Clear
    For i = 0 To cat.Tables.Count - 1
        If LCase(cat.Tables(i).Type) = "table" Then
            Form1.List1.AddItem cat.Tables(i).Name
        End If
    Next
    Form2.Hide
errH:
    Dim pw As String
    'check if there is a database password
    If Err.Number = -2147217843 Then
        Text1.Text = ""
        MsgBox "Invalid password!"
        Text1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
    Form2.Hide
End Sub
