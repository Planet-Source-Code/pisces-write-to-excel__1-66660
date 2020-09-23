VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2805
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtSaveTo 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "C:\Text.xls"
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Cell 2"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Cell 1"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Save To:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ExlObj As Excel.Application
Dim ExlWK As Excel.Workbook
Dim ExlWS As Excel.Worksheet


Private Sub Command1_Click()
 Set ExlObj = CreateObject("excel.application")      ' Initialize the excel object
    Set ExlWK = ExlObj.Workbooks.Add
    Set ExlWS = ExlWK.Worksheets(1)
  
  'ExlObj.Visible = True ' So you can see Excel
  ExlObj.Range("A1:Z1").Borders.Color = RGB(1, 3, 7) 'Use it to change the borders.
  ExlObj.Columns("A:AY").EntireColumn.AutoFit 'To adjust the column's width.
    
    ' Print the Info
    With ExlWS
        .Cells(1, 1).Value = Text1.Text
        .Cells(1, 2).Value = Text2.Text
        'Might add on data into cell like .Cell(2,3).... so on
    End With
        
  Set ExlWS = Nothing
    ExlWK.SaveAs txtSaveTo.Text
    ExlWK.Close
    MsgBox "File Created Successfully !"
   
End Sub


