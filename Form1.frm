VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0634
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Viewer 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5530
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iHeader As String
Dim xFile As String
Dim iName As String
Dim iData As String

Private Sub Form_Load()

Dim TempString As String
Dim xNode As Node
Dim i As Integer

    xFile = Command$

    If Command$ = "" Then xFile = App.Path & "\system.ini"

    With Viewer
        .Nodes.Clear

        Set xNode = .Nodes.Add(, , "A", xFile, 3)

    Open xFile For Input As #1
        
        Do While Not EOF(1)

            Line Input #1, TempString
                           
                If TempString = "" Then GoTo InputIt
                        
                If Mid(TempString, 1, 1) = "[" Then
                    i = i + 1
                    Set xNode = .Nodes.Add("A", tvwChild, "A" & i, TempString, 1)
                    GoTo InputIt
                End If

             Set xNode = .Nodes.Add("A" & i, tvwChild, , TempString, 2)
InputIt:
        Loop

    End With
        
End Sub

Private Sub Form_Resize()

Viewer.Move 1, 1, Form1.ScaleWidth, Me.ScaleHeight

End Sub

Private Sub Viewer_AfterLabelEdit(Cancel As Integer, NewString As String)

iName = Split(NewString, "=")(0)
iData = Split(NewString, "=")(1)
iHeader = Mid(Viewer.SelectedItem.Parent.Text, 2, Len(Viewer.SelectedItem.Parent.Text) - 2)

MsgBox "[" & iHeader & "]" & " ( " & iName & " ) " & "= " & iData

WriteINI iHeader, iName, iData, xFile

End Sub

Private Sub Viewer_Click()

If Mid(Viewer.SelectedItem.Text, 1, 1) = "[" Or Viewer.SelectedItem.Index = 1 Then Exit Sub
Viewer.StartLabelEdit

End Sub
