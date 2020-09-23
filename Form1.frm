VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Colorizer"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   4275
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "Form1.frx":014A
      Top             =   0
      Width           =   6165
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   690
      Left            =   900
      TabIndex        =   7
      Top             =   300
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "*"
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change"
      Enabled         =   0   'False
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   240
      Left            =   1800
      TabIndex        =   2
      Top             =   750
      Width           =   1215
   End
   Begin VB.OptionButton optSize 
      Caption         =   "140  * 42"
      Height          =   240
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Tag             =   "70         42"
      Top             =   225
      Width           =   1140
   End
   Begin VB.OptionButton optSize 
      Caption         =   "70 * 42"
      Height          =   240
      Index           =   2
      Left            =   1800
      TabIndex        =   0
      Tag             =   "70         42"
      Top             =   450
      Value           =   -1  'True
      Width           =   1140
   End
   Begin RichTextLib.RichTextBox rich1 
      Height          =   9015
      Left            =   -75
      TabIndex        =   6
      Top             =   1125
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   15901
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton optSize 
      Caption         =   "70 * 42"
      Height          =   240
      Index           =   0
      Left            =   1800
      TabIndex        =   5
      Tag             =   "70         42"
      Top             =   225
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblY 
      Caption         =   "0"
      Height          =   165
      Left            =   3525
      TabIndex        =   13
      Top             =   675
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblX 
      Caption         =   "0"
      Height          =   165
      Left            =   3525
      TabIndex        =   12
      Top             =   675
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Label lblMove 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                                T H I S   I S   T H E  R E S I Z E  B A R"
      Height          =   8940
      Left            =   4275
      MousePointer    =   9  'Size W E
      TabIndex        =   11
      Top             =   1125
      Width           =   165
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblColor 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   690
      Left            =   0
      TabIndex        =   10
      Top             =   300
      Width           =   915
   End
   Begin VB.Image imgOrig 
      BorderStyle     =   1  'Fixed Single
      Height          =   8925
      Left            =   5700
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "Form1.frx":0E66
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "( | symbol)"
      Height          =   240
      Left            =   3000
      TabIndex        =   9
      Top             =   225
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "( * symbol)"
      Height          =   240
      Left            =   3000
      TabIndex        =   8
      Top             =   450
      Width           =   765
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************************************
'********Made by Rob Heineman*****************

Private Sub Command1_Click()
    'just some error trapping because it hardly ever works
    On Error GoTo hi
    
    Me.Top = 0
    Me.Left = 0
    Dim t, d, moveInt, moveInty As Integer
        'Find Intervals to draw from
        moveInt = Round(imgOrig.Width / Mid(optSize(0).Tag, 1, 3), 3)
        moveInty = Round(imgOrig.Height / Right(optSize(0).Tag, 3), 3)
        For d = 0 To Right(optSize(0).Tag, 3)
            For t = 0 To Mid(optSize(0).Tag, 1, 3)
                lblColor.BackColor = Point(imgOrig.Left + (t * moveInt), imgOrig.Top + (d * moveInty))
                rich1.SelStart = (d * (Mid(optSize(0).Tag, 1, 3) - 1)) + t
                rich1.SelLength = 1
                rich1.SelColor = lblColor.BackColor
            Next t
        Next d
        GoTo Skip  'Ok you guys can kill me for using goto functions
hi:
    MsgBox "An error occurred make sure that the form is topmost and the image is not obstructed.  I'm sorry I wrote such a tempremental program"
Skip:
End Sub

Private Sub Command2_Click()
    Dim t, d As Integer
    'Cycles through and resets the textbox
    If Len(Text1) > 0 Then
            For t = 0 To Len(rich1.Text)
                d = d + 1
                If d > Len(Text1) Then
                    d = 1
                End If
                rich1.SelStart = t
                rich1.SelLength = 1
                rich1.SelText = Mid(Text1, d, 1)
            Next t
    End If
End Sub

Private Sub Command3_Click()
    rich1.Text = "" 'Clears Box
End Sub

Private Sub lblMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This is a peice of moving code I use on all my projects
    If Button > 0 Then
        Dim Xmove As Integer
        If lblX = 0 Then
            lblX = X - 30
        End If
        Xmove = X - lblX
        With rich1
            'Resize Control
            .Width = .Width + Xmove
            'Move Resize Bar
            lblMove.Left = lblMove.Left + Xmove
        End With
    End If
End Sub

Private Sub optSize_Click(Index As Integer)
    If Index = 1 Then
        Text1 = "|"
    End If
    If Index = 2 Then
        Text1 = "*"
    End If
    'Sets the tag to the Options Caption
    optSize(0).Tag = optSize(Index).Caption
    'Clear Texbox
    rich1.Text = ""
    'Refills Textbox
    Dim t, d As Integer
    For d = 0 To Right(optSize(0).Tag, 3)
        For t = 0 To Mid(optSize(0).Tag, 1, 3)
            'This is just a filler the text box just needs to be x characters long
            rich1.Text = rich1.Text + "/"
        Next t
    Next d
    'Fills with the desired character
    'I was going to design it differently and thats why its somewhere else
    Call Command2_Click
End Sub
