VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Type to search the listbox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      X1              =   120
      X2              =   5520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Type to search the combobox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As String, lParam As Any) As Long
Const LB_FINDSTRING = &H18F


Private Sub Form_Load()
With List1
    .AddItem "apple"
    .AddItem "ant"
    .AddItem "bird"
    .AddItem "blue"
    .AddItem "car"
    .AddItem "cars"
    .AddItem "Deep"
    .AddItem "deeper"
    .AddItem "Garage"
    .AddItem "greed"
    .AddItem "jeep"
    .AddItem "Jail"
    .AddItem "monitor"
    .AddItem "melt"
    .AddItem "rest"
    .AddItem "risk"
End With

With Combo1
    .AddItem "apple"
    .AddItem "ant"
    .AddItem "bird"
    .AddItem "blue"
    .AddItem "car"
    .AddItem "cars"
    .AddItem "Deep"
    .AddItem "deeper"
    .AddItem "Garage"
    .AddItem "greed"
    .AddItem "jeep"
    .AddItem "Jail"
    .AddItem "monitor"
    .AddItem "melt"
    .AddItem "rest"
    .AddItem "risk"
End With

End Sub


Private Sub Text1_Change()

Dim i As Long

For i = 0 To Combo1.ListCount - 1
Combo1.ListIndex = i
If LCase(Left(Combo1.Text, Len(Text1.Text))) = LCase(Text1.Text) Then
Exit For
End If
Next i

If LCase(Left(Combo1.Text, Len(Text1.Text))) <> LCase(Text1.Text) Then
SendKeys "{backspace}"
End If

If Text1.Text = "" Then Combo1.ListIndex = -1

End Sub


Private Sub Text2_Change()
List1.ListIndex = SendMessage(List1.hWnd, LB_FINDSTRING, Text2, ByVal Text2.Text)

End Sub


