VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFast 
   Caption         =   "Fast Load"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   7005
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   5760
      Width           =   3615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5175
      Left            =   4320
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Using this method fastly return all rows, we also still have recordset as cache. so we can search for record without extra load."
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   6240
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   4335
   End
End
Attribute VB_Name = "frmFast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fRs As ADODB.Recordset
Private fConn As ADODB.Connection
Private fTime As Date

Private Sub Form_Load()
    fTime = Now
    Set fConn = New ADODB.Connection
    fConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
    Set fRs = New ADODB.Recordset
    'limit recordset page size to 20
    fRs.PageSize = 20
    fRs.Open "select * from inventory order by code", fConn, adOpenStatic, adLockOptimistic
    'set grid rows to 21
    grid1.Rows = 21
    'set vscroll min and max and display 20 first records
    VScroll1.Min = 1
    VScroll1.Max = fRs.PageCount
    DoEvents
    Label1.Caption = "Time : " & Str(DateDiff("s", fTime, Now)) & " Second"
End Sub

Private Sub Display()
    Dim i As Long
    Dim x As Long
    
    'clear grid first
    'because last page may less than 20 records
    grid1.Clear
    ' Set column names in the grid
    For i = 0 To fRs.Fields.Count - 1
        grid1.TextMatrix(0, i) = fRs.Fields(i).Name
    Next
    grid1.Redraw = False
    'display 20 record only
    i = 1
    Do While (i <= 20) And (Not fRs.EOF)
        For x = 0 To fRs.Fields.Count - 1
            Me.grid1.TextMatrix(i, x) = fRs.Fields(x).Value
        Next
        fRs.MoveNext
        i = i + 1
    Loop
    grid1.Redraw = True
End Sub

Private Sub Text1_Change()
    If Text1.Text = "" Then
        fRs.MoveFirst
    Else
        fRs.Find "[" & fRs.Fields(1).Name & "] >= '" & Text1.Text & "'", 0, adSearchForward, 1
    End If
    If Not fRs.EOF Then
        VScroll1.Value = fRs.AbsolutePage
    End If
End Sub

Private Sub VScroll1_Change()
    fRs.AbsolutePage = VScroll1.Value
    Display
End Sub
