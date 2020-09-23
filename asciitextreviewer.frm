VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ascii Text Reviewer - By Phishbowler"
   ClientHeight    =   2835
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5385
   Icon            =   "asciitextreviewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   5385
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2205
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1440
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00004000&
      Caption         =   "View Ascii Character Chart"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   5415
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0000FF00&
      Caption         =   "Type Text"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0000FF00&
      Caption         =   "Paste Text"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4680
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Chr$ Code"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FF00&
      Caption         =   "Chr$ Count"
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Code Was Written By: Phishbowler
'Sept. 14, 2000
'
'For Money Making Opportunities, Visit
'Http://www.dreamstruct.com/
'
'Napster Users: Tired of Incomplete Songs?
'Get the good ol' Nap v2.0 Only available at:
'Http://come.to/NapsterResume
'
'The color form fade was
'written by the same author
'who wrote cryofade.bas
'for AIM




Private Sub Command1_Click()
If Text1 = "" Then Exit Sub
Do
a = a + 1
B = Mid(Form1.Text1, 1, a)
Loop Until a = Len(Form1.Text1)
Form1.Text3 = a
End Sub

Private Sub Check1_Click()

If Check1.Value = 1 Then
Form1.Height = 7740
Else
Form1.Height = 3240
End If
End Sub

Private Sub Form_DblClick()
Form1.WindowState = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Static LastAmount As Long

AmountList = List1.ListCount

    'If Backspace Detected
If KeyAscii = 8 Then
    'If CHRTextEnter has text and 1 chr for text1 then clear!
If Form1.Text4 > "" And Len(Form1.Text1) = 1 Then
Form1.Text4 = ""
Form1.Text1 = ""
Form1.Option1.Value = True
End If

'Check to see if they typed in text4
If Len(Form1.Text4) > 0 Then Exit Sub
Else

'If CNTRL-V (Paste) Detected Then..
If KeyAscii = 22 Then
Form1.List1.Clear
Form1.Option2.Value = True
Form1.Option1.Value = False
Form1.Text1.SetFocus
GoTo PasteText:
End If

'Type Text Option
If Form1.Option1.Value = True Then
If KeyAscii = 0 Then GoTo Skip
newchar = Chr$(KeyAscii)
If newchar = "" Then GoTo Skip
List1.AddItem newchar & " - " & KeyAscii: Exit Sub
End If

Skip:


PasteText:
If KeyAscii = 0 And Len(Form1.Text1) < LastAmount Then Exit Sub
If Form1.Option2.Value = True Then
LastAmount = 0
LastAmount = Len(Form1.Text1)
 If Text1 = "" Then Exit Sub
 
'Go Through Pasted Text and Add To List
 Do
 a = a + 1
 source = Mid$(Form1.Text1, a, 1)
 For X = 0 To 255
  If source = Chr$(X) Then
  List1.AddItem source & " - " & X
  End If
  Next X
  
  
  Loop Until a = Len(Form1.Text1)
  End If
Form1.Text3 = Len(Form1.Text1)
If Form1.Text1 = "" Then List1.Clear
Done:
End If

End Sub

Private Sub Form_Load()
'Place Window On Top
Call Win_OnTop(Form1)

'This is the ASCII Chart at the bottom
For X = 0 To 255
Form1.Text2 = Form1.Text2 & X & " " & Chr$(X) & " "
Next X

End Sub

Private Sub Form_Paint()
'All FormFade's in GeneralAPI Credit is due to
'writer of CRYOFADE.BAS for AIM
Call FormFadeGreen(Form1)
End Sub

Private Sub Form_Resize()
If Form1.WindowState = 2 Then Form1.WindowState = 0
If Form1.WindowState = 0 Then
If Form1.Width <> 5505 Then Form1.Width = 5505
If Form1.Height > 7740 Then Form1.Height = 7740
End If
End Sub

Private Sub Label1_DblClick()
Form1.WindowState = 1
End Sub

Private Sub Label2_DblClick()
Form1.WindowState = 1
End Sub

Private Sub List1_DblClick()
Form1.WindowState = 1
End Sub

Private Sub Option1_Click()
Form1.List1.Clear
Form1.Text1 = ""
Form1.Text4 = ""
Form1.Text1.SetFocus
End Sub

Private Sub Option2_Click()
Form1.List1.Clear
Form1.Text1 = ""
Form1.Text4 = ""
Form1.Text1.SetFocus
End Sub

Private Sub Option2_DblClick()
Form1.WindowState = 1
End Sub

Private Sub Text1_Change()
Static LastAmount As Integer
Static charcountold As Integer
On Error GoTo ErrorHandle:
If Text3 = "" Then Text3 = 0
Form_KeyPress (0)
If Len(Text1) = 0 And List1.ListCount = 0 Then Exit Sub
If Len(Text1) > 1 And Len(Text4) > 0 Then Text4 = "": Text1 = "": Form1.Option1.Value = True

If Len(Text4) > 0 Then GoTo CarriageDel:
If Text4 = "" And Text1 = "" Then GoTo CarriageDel:
If charcountold = 0 Then charcountold = Len(Text1): Exit Sub
If Len(Form1.Text1) < charcountold Then
LastAmount = Len(Form1.Text1)
AmountList = Form1.List1.ListCount
If List1.List(AmountList - 1) = Chr$(10) & " - 10" Then
Form1.List1.RemoveItem (List1.ListCount - 1)
Form1.List1.RemoveItem (List1.ListCount - 1)
charcountold = charcountold - 2

End If

If charcountold > Len(Form1.Text1) Then

Difference = charcountold - Len(Form1.Text1)
For X = 1 To Difference
List1.RemoveItem (List1.ListCount - 1)
Next X
End If

End If
CarriageDel:

charcountold = Len(Form1.Text1)



Form1.Text3 = Len(Form1.Text1)
If Form1.Text1 = "" Then List1.Clear

'I wasn't sure how to detect when someone is typing
'when in paste mode, so I use an error handler,
'If you feel like fixing be my guest.

ErrorHandle:

Select Case Err.Number
Case 0:
If List1.List(0) = Form1.List1.List(1) And Form1.Option2.Value = True Then List1.Clear: Text3 = 0: Text1 = ""

End Select
End Sub


Private Sub Text1_DblClick()
Form1.WindowState = 1
End Sub

Private Sub Text2_DblClick()
Form1.WindowState = 1
End Sub

Private Sub Text3_DblClick()
Form1.WindowState = 1
End Sub

Private Sub Text4_Change()
On Error GoTo sucks
Dim NumCheck As Integer


List1.Clear
NumCheck = Text4


If NumCheck > 255 Then Text4 = "": Exit Sub
If NumCheck < 0 Then Text4 = "": Exit Sub
If Text4 = "" Then: Text1 = "": Exit Sub
Form1.Text1 = Chr$(Text4)

'This just clears Chr$Code box if they delete the last character
If Text1 = "" And Text4 > "" And Text3 = "0" Then
Text4 = ""
Form1.Option1.Value = True
End If


sucks:
Select Case Err.Number
Case 13:
Text4 = ""
Text1 = ""

End Select

End Sub

Private Sub Text4_Click()
Form1.List1.Clear
Text1.Text = ""
Form1.Option1.Value = False
Form1.Option2.Value = False
End Sub

Private Sub Text4_DblClick()
Form1.WindowState = 1
End Sub
