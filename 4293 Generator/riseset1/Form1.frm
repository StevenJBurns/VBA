VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Global sun & moon rise/set calculator"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3645
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   23
      Top             =   1905
      Width           =   240
   End
   Begin VB.TextBox Text10 
      Height          =   345
      Left            =   4680
      TabIndex        =   9
      Top             =   1680
      Width           =   1065
   End
   Begin VB.TextBox Text9 
      Height          =   345
      Left            =   4680
      TabIndex        =   10
      Top             =   2055
      Width           =   1065
   End
   Begin VB.TextBox Text8 
      Height          =   345
      Left            =   4680
      TabIndex        =   8
      Top             =   1260
      Width           =   1065
   End
   Begin VB.TextBox Text7 
      Height          =   345
      Left            =   4680
      TabIndex        =   7
      Top             =   885
      Width           =   1065
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3645
      Picture         =   "Form1.frx":0B14
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   1110
      Width           =   240
   End
   Begin VB.TextBox Text6 
      Height          =   345
      Left            =   4680
      TabIndex        =   5
      Top             =   90
      Width           =   1065
   End
   Begin VB.TextBox Text5 
      Height          =   345
      Left            =   4680
      TabIndex        =   6
      Top             =   465
      Width           =   1065
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   150
      Picture         =   "Form1.frx":0D5E
      ScaleHeight     =   180
      ScaleWidth      =   4185
      TabIndex        =   16
      Top             =   2595
      Width           =   4185
   End
   Begin VB.TextBox Text4 
      Height          =   345
      Left            =   1005
      TabIndex        =   3
      Top             =   1650
      Width           =   750
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   1005
      TabIndex        =   2
      Top             =   1275
      Width           =   1200
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   1005
      TabIndex        =   1
      Top             =   900
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1005
      TabIndex        =   0
      Top             =   525
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4605
      TabIndex        =   11
      Top             =   2505
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate ->"
      Default         =   -1  'True
      Height          =   375
      Left            =   2340
      Picture         =   "Form1.frx":2CB8
      TabIndex        =   4
      Top             =   1035
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Nautical twilight"
      Height          =   480
      Left            =   3930
      TabIndex        =   22
      Top             =   1830
      Width           =   870
   End
   Begin VB.Label Label8 
      Caption         =   "Moonset"
      Height          =   225
      Left            =   3930
      TabIndex        =   21
      Top             =   1305
      Width           =   1035
   End
   Begin VB.Label Label7 
      Caption         =   "Moonrise"
      Height          =   255
      Left            =   3930
      TabIndex        =   20
      Top             =   945
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   195
      Left            =   3645
      Picture         =   "Form1.frx":3582
      Top             =   300
      Width           =   195
   End
   Begin VB.Label Label6 
      Caption         =   "Sunrise"
      Height          =   255
      Left            =   3930
      TabIndex        =   18
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Sunset"
      Height          =   225
      Left            =   3930
      TabIndex        =   17
      Top             =   510
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Timezone"
      Height          =   300
      Left            =   180
      TabIndex        =   15
      Top             =   1710
      Width           =   1110
   End
   Begin VB.Label Label3 
      Caption         =   "Longitude"
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   1335
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Latitude"
      Height          =   225
      Left            =   180
      TabIndex        =   13
      Top             =   945
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   585
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
' Riseset.DLL must be registered for this program to work!
Dim RS As RiseSetDLL.Class
Set RS = New RiseSetDLL.Class
valid = True
If Not IsDate(Text1) Then
    MsgBox ("Please provide a valid date")
    valid = False
End If
If Not (IsNumeric(Text2) And IsNumeric(Text3)) Then
    MsgBox ("Please enter a numeric latitude and longitude")
    valid = False
End If
If Not IsNumeric(Text4) Then
    MsgBox ("Please enter a numeric timezone between -12 and +12")
    valid = False
Else
    If Text4 < -12 Or Text4 > 12 Then
        MsgBox ("Please enter a numeric timezone between -12 and +12")
        valid = False
    End If
End If
If valid Then
    Text6 = RS.SRS(Text2, Text3, Text4, "R", Year(Text1), Month(Text1), Day(Text1))
    Text5 = RS.SRS(Text2, Text3, Text4, "S", Year(Text1), Month(Text1), Day(Text1))
    Text7 = RS.SRS(Text2, Text3, Text4, "MR", Year(Text1), Month(Text1), Day(Text1))
    Text8 = RS.SRS(Text2, Text3, Text4, "MS", Year(Text1), Month(Text1), Day(Text1))
    Text10 = RS.SRS(Text2, Text3, Text4, "NR", Year(Text1), Month(Text1), Day(Text1))
    Text9 = RS.SRS(Text2, Text3, Text4, "NS", Year(Text1), Month(Text1), Day(Text1))
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
On Error GoTo nodll
Dim RS As RiseSetDLL.Class
Text1 = Format(Now, "Short Date")
Set RS = New RiseSetDLL.Class
GoTo endsub
nodll:
    Me.Visible = True
    If MsgBox("The file riseset.dll needs to be registered for this program to run.  Is it ok to register it now?", vbDefaultButton2 + vbQuestion + vbYesNo, "RiseSet.DLL not registered!") = vbYes Then x = Shell("regsvr32 " & App.Path & "\riseset.dll", vbNormalFocus) Else End
endsub:
End Sub
