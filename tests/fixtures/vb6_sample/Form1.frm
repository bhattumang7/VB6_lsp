VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSChrt20.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2655
      Left            =   600
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   1
      Top             =   3600
      Width           =   3735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2655
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4683
      _Version        =   393216
      Rows            =   8
      Cols            =   4
      MouseIcon       =   "Form1.frx":2356
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
