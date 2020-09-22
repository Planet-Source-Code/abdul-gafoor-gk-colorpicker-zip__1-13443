VERSION 5.00
Object = "*\AClrPckr.vbp"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test ColorPicker Control"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Show ToolTips"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Width           =   1455
   End
   Begin ClrPckr.ColorPicker ColorPicker1 
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "Flat"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   10
      Top             =   240
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      Caption         =   "3D"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   9
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Show More Colors Button"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Show System Color Button"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Show Custom Colors"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      Caption         =   "Show Default Color"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "More Colors Caption"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Default Caption"
      Height          =   195
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click(Index As Integer)
    Select Case Index
        Case 0: If Check1(0).Value = 1 Then ColorPicker1.ShowDefault = True Else ColorPicker1.ShowDefault = False
        Case 1: ColorPicker1.ShowCustomColors = Check1(Index).Value
        Case 2: ColorPicker1.ShowSysColorButton = Check1(Index).Value
        Case 3: ColorPicker1.ShowMoreColors = Check1(Index).Value
        Case 4: ColorPicker1.ShowToolTips = Check1(Index).Value
    End Select
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    With Me.ColorPicker1
        For i = 0 To Check1.Count - 1
            Check1(i).Value = 1
        Next i
        
        If .Appearance = [3D] Then
            Option1(0).Value = True
        Else
            Option1(1).Value = True
        End If
        
        Text1(0) = .DefaultCaption
        Text1(1) = .MoreColorsCaption
    End With
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0: ColorPicker1.Appearance = [3D]
        Case 1: ColorPicker1.Appearance = Flat
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Select Case Index
        Case 0: ColorPicker1.DefaultCaption = Text1(Index)
        Case 1: ColorPicker1.MoreColorsCaption = Text1(Index)
    End Select
End Sub
