VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Notepad++"
   ClientHeight    =   6780
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   12480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0666
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0778
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":088A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":099C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BC0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Status Bar"
      Top             =   6405
      Visible         =   0   'False
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
            Text            =   "Press F1 for help"
            TextSave        =   "Press F1 for help"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "time"
            Object.ToolTipText     =   "Time/Date"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdg 
      Left            =   1080
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtb 
      Height          =   3375
      HelpContextID   =   1
      Left            =   840
      TabIndex        =   0
      ToolTipText     =   "Text Box"
      Top             =   1560
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   5953
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":1012
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu saveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu print 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu time 
         Caption         =   "Time&Date"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu format 
      Caption         =   "&Format"
      Begin VB.Menu font 
         Caption         =   "Font"
      End
      Begin VB.Menu color 
         Caption         =   "Color"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu status 
         Caption         =   "Status Bar"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub color_Click()
cmdg.ShowColor
rtb.SelColor = cmdg.color
End Sub

Private Sub Command1_Click()
Dim n As Integer
n = InputBox("What to find?", vbOKCancel)

End Sub



Private Sub copy_Click()
Clipboard.Clear
Clipboard.SetText rtb.SelText, vbCFText
End Sub

Private Sub cut_Click()
Clipboard.Clear
Clipboard.SetText rtb.SelText, vbCFText
rtb.SelText = ""
End Sub

Private Sub exit_Click()
Dim n As Integer
n = MsgBox("Are you Sure?", vbYesNo + vbQuestion, "Exit")
If n = vbYes Then
End
End If
End Sub



Private Sub font_Click()
cmdg.Flags = &H3
cmdg.ShowFont
rtb.SelFontName = cmdg.FontName
rtb.SelFontSize = cmdg.FontSize
rtb.SelBold = cmdg.FontBold
rtb.SelUnderline = cmdg.FontUnderline
rtb.SelItalic = cmdg.FontItalic
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbar.Panels(2) = X & " , " & Y
End Sub

Private Sub Form_Resize()
rtb.Left = 0
rtb.Top = tb.Height + 100
rtb.Height = Form1.ScaleHeight
rtb.Width = Form1.ScaleWidth
End Sub

Private Sub help_Click()
Dim s As Integer
s = MsgBox("Please check help registry", vbOKOnly, "Error")
End Sub

Private Sub new_Click()
rtb.Text = " "
End Sub

Private Sub open_Click()
cmdg.ShowOpen
rtb.LoadFile cmdg.FileName
End Sub

Private Sub paste_Click()
rtb.SelText = Clipboard.GetText(vbCFText)
End Sub

Private Sub print_Click()
cmdg.Flags = cdlCFBoth
cmdg.ShowPrinter
End Sub

Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sbar.Panels(2) = X & " , " & Y
End Sub

Private Sub save_Click()
cmdg.DefaultExt = "txt"
cmdg.Filter = "txt files|*.txt"
cmdg.ShowSave
rtb.SaveFile cmdg.FileName, txtTXT
cmdg.Flags = cdlOFNOverwritePrompt
End Sub

Private Sub saveas_Click()
cmdg.DefaultExt = "txt"
cmdg.Filter = "txt files|*.txt"
cmdg.ShowSave
rtb.SaveFile cmdg.FileName, txtTXT
End Sub





Private Sub status_Click()
sbar.Visible = True
End Sub

Private Sub tb_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
rtb.Text = " "
ElseIf Button.Index = 2 Then
cmdg.ShowOpen
rtb.LoadFile cmdg.FileName
ElseIf Button.Index = 3 Then
cmdg.DefaultExt = "txt"
cmdg.Filter = "txt files|*.txt"
cmdg.ShowSave
rtb.SaveFile cmdg.FileName, txtTXT
cmdg.Flags = cdlOFNOverwritePrompt
ElseIf Button.Index = 4 Then
cmdg.Flags = cdlCFBoth
cmdg.ShowPrinter
ElseIf Button.Index = 5 Then
Clipboard.Clear
Clipboard.SetText rtb.SelText, vbCFText
rtb.SelText = ""
ElseIf Button.Index = 6 Then
Clipboard.Clear
Clipboard.SetText rtb.SelText, vbCFText
ElseIf Button.Index = 7 Then
rtb.SelText = Clipboard.GetText(vbCFText)
ElseIf Button.Index = 8 Then
rtb.SelText = Now
End If
End Sub

Private Sub time_Click()
rtb.SelText = Now
End Sub

