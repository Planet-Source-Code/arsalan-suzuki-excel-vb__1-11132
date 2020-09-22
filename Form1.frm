VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Excel Demo Program"
   ClientHeight    =   1155
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Go on, Just Click the button and u wil get the result"
      ForeColor       =   &H00FF0000&
      Height          =   990
      Left            =   150
      TabIndex        =   0
      Top             =   45
      Width           =   4050
      Begin VB.CommandButton cmdStart 
         Caption         =   "Run Microsoft Excel"
         Default         =   -1  'True
         Height          =   495
         Left            =   180
         TabIndex        =   1
         Top             =   315
         Width           =   3690
      End
   End
   Begin VB.Menu mnuAuthor 
      Caption         =   "About The Author"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/-------About this program---------\
'Well I made this program just to
'demonstrate the way to run excel from
'Visual Basic. Write few lines in excel,
'format it, then save the file.
'When I saw that I couldnt find this type
'of program I wrote this one , so that people
'could learn something new.
'If you want to contact me address can be
'found on about form
'-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'I'm sure that you will learn something
'new from it
'-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'If you have problem be sure to check
'1.(Menu) Project----->References
'2.Check Miscrosoft Excel 9.0 Object Library
'If you have any further problem
'contact me
'I have tested this in Excel 2000 so I dont
'know whether it will work in prev version
'of Excel
'\----------------------------------/

'Define the required variable
Dim Excel As Excel.Application ' This is the excel program
Dim ExcelWBk As Excel.Workbook ' This is the work book
Dim ExcelWS As Excel.Worksheet ' This is the sheet

'Well I have broken this program in to several subs.
'This is the main sub from where every thing will
'be called.
Private Sub cmdStart_Click()
If cmdStart.Caption = "Run Microsoft Excel" Then
    StartExcel
    cmdStart.Caption = "Create WorkSheet"
    Exit Sub ' Otherwise it will do everything in one shot
ElseIf cmdStart.Caption = "Create WorkSheet" Then
    CreateWorkSheet
    cmdStart.Caption = "Populate WorkSheet"
    Exit Sub
ElseIf cmdStart.Caption = "Populate WorkSheet" Then
    PopulateWorkSheet
    cmdStart.Caption = "Format The Sheet"
    Exit Sub
ElseIf cmdStart.Caption = "Populate WorkSheet" Then
    PopulateWorkSheet
    cmdStart.Caption = "Format The Sheet"
    Exit Sub
ElseIf cmdStart.Caption = "Format The Sheet" Then
    FormatWorkSheet
    cmdStart.Caption = "Save The WorkBook"
    Exit Sub
ElseIf cmdStart.Caption = "Save The WorkBook" Then
    SaveWorkSheet
    cmdStart.Caption = "Close The WorkBook and Excel"
    Exit Sub
ElseIf cmdStart.Caption = "Close The WorkBook and Excel" Then
    CloseWorkSheet
    cmdStart.Caption = "Bye , My demonstration is finished."
    Exit Sub
ElseIf cmdStart.Caption = "Bye , My demonstration is finished." Then
    Unload Me
End If

End Sub

Private Sub mnuAuthor_Click()
frmAuthor.Show 1
End Sub

Private Sub StartExcel()
On Error GoTo err:

Set Excel = GetObject(, "Excel.Application") ' Create Excel Object.
'Well you have to do like this.
'Above line if I used CreateObject, 1st time it would
'work fine but the second time my program would
'hang.Well I found this the easiest way to do it.
'But you can do it another way if you like.


'By default after creating the Excel it will
'not be shown on the screen.
'We will show it later.....
'Excel.Visible = True ' Show Excel

Exit Sub
err:
Set Excel = CreateObject("Excel.Application") 'Create Excel Object.

End Sub

Private Sub CreateWorkSheet()
Set ExcelWBk = Excel.Workbooks.Add 'Add this Workbook to Excel.
Set ExcelWS = ExcelWBk.Worksheets(1) ' Add this sheet to this Workbook

End Sub

Private Sub PopulateWorkSheet()
Dim col As Integer
Dim row As Integer
Randomize Timer ' Random, To generate random number.....

For col = 1 To 5 ' coloumn
   For row = 1 To 20 ' row
      ExcelWS.Cells(row, col) = Rnd() * 100 'Populate with random runmber
   Next row
Next col

End Sub

Private Sub FormatWorkSheet()
'Format the cell from A1 to E20
ExcelWS.Range("A1:E20").NumberFormat = "0.00"
End Sub

Private Sub SaveWorkSheet()
' Save the workbook on the desktop
ExcelWBk.SaveAs "c:\windows\desktop\Demo.xls"
End Sub

Private Sub CloseWorkSheet()
' Close the WorkBook
ExcelWBk.Close
' Quit Excel app
Excel.Quit

MsgBox "You can find the saved Excel Sheet on your desktop"

End Sub
