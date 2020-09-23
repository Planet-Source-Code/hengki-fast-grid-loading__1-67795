VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBound 
   Caption         =   "Bound Grid To Recordset"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   6225
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
      Caption         =   "This Method only return +/- 2000 rows"
      Height          =   375
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
Attribute VB_Name = "frmBound"
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
    fRs.Open "select * from inventory order by code", fConn, adOpenStatic, adLockOptimistic
    Set grid1.DataSource = fRs
    DoEvents
    Label1.Caption = "Time : " & Str(DateDiff("s", fTime, Now)) & " Second"
End Sub
