VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmClip 
   Caption         =   "Load Using Clip Function"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   6630
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      Caption         =   $"frmClip.frx":0000
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   5880
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
Attribute VB_Name = "frmClip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fRs As ADODB.Recordset
Private fConn As ADODB.Connection
Private fTime As Date

Private Sub Form_Load()
    Dim rsVar As Variant
    Dim i As Long
    
    fTime = Now
    Set fConn = New ADODB.Connection
    fConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
    Set fRs = New ADODB.Recordset
    fRs.Open "select * from inventory order by code", fConn, adOpenStatic, adLockOptimistic
    grid1.Rows = fRs.RecordCount + 1
    rsVar = fRs.GetString(adClipString, fRs.RecordCount)
    ' Set column names in the grid
    For i = 0 To fRs.Fields.Count - 1
        grid1.TextMatrix(0, i) = fRs.Fields(i).Name
    Next
    grid1.Row = 1
    grid1.Col = 0
    ' Set range of cells in the grid
    grid1.RowSel = grid1.Rows - 1
    grid1.ColSel = grid1.Cols - 1
    grid1.Clip = rsVar
    ' Reset the grid's selected range of cells
    grid1.RowSel = grid1.Row
    grid1.ColSel = grid1.Col
    DoEvents
    Label1.Caption = "Time : " & Str(DateDiff("s", fTime, Now)) & " Second"
End Sub
