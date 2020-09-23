VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fast Grid Loading"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Fast Load"
      Height          =   495
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Using Clip Function"
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bound Grid To Recordset"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        frmBound.Show
    ElseIf Index = 1 Then
        frmClip.Show
    Else
        frmFast.Show
    End If
End Sub

Private Sub Form_Load()
    Dim pConn As ADODB.Connection
    Dim pRs As ADODB.Recordset
    Dim i As Long
    
    'this is to append data to inventory table
    'i do this because i dont want to upload large file
    'my internet connection is slow :(
    'so i append it when you show this form for first time
    
    Set pConn = New ADODB.Connection
    pConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
    Set pRs = New ADODB.Recordset
    pRs.Open "select count(code) as reccount from inventory", pConn, adOpenForwardOnly, adLockReadOnly
    If pRs!recCount < 2 Then
        For i = 1 To 16
            pConn.Execute "INSERT INTO inventory ( Code, Name ) " & _
                        "SELECT Query2.expr1, Query2.expr2 " & _
                        "FROM Query2", adExecuteNoRecords
        Next
    End If
End Sub
