VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort Example"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwTest 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5318
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgAll 
      Left            =   5400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sortlistview.frx":0000
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sortlistview.frx":0452
            Key             =   "down"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "sortlistview.frx":08A4
            Key             =   "no"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblClick 
      Caption         =   "&Click on column headers to see the sort ... ( especially on text numeric and text simple)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    AddColumns
    AddData
    'associate with imagelist - optional
    lvwTest.ColumnHeaderIcons = imgAll
End Sub
Private Sub AddData()
    Dim i As Long
    For i = 1 To 100
        With lvwTest.ListItems.Add
            .Text = i
            .ListSubItems.Add , , i
            .ListSubItems.Add , , Date - i
        End With
    Next i
End Sub
Private Sub AddColumns()
    ' for use of column headers tags see SetListViewOrder
    
    With lvwTest.ColumnHeaders
        With .Add
            .Text = "Text Numeric"
            .Tag = "numeric" ' this tag helps to identify numeric columns
            
        End With
        With .Add
            .Text = "text simple"
            .Tag = "" 'modify here in numeric
        End With
        With .Add
            .Text = "date"
            .Tag = "date" ' this tag helps to identify numeric columns
        End With
        
    End With
    
End Sub

Private Sub lvwTest_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    SetListViewOrder lvwTest, ColumnHeader
End Sub
