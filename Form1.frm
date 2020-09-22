VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ASP Parser (technical@cemhaner.com)"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Receive data"
      Height          =   570
      Left            =   2070
      TabIndex        =   11
      Top             =   3765
      Width           =   2115
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   1515
      TabIndex        =   10
      Top             =   2205
      Width           =   3630
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   1515
      TabIndex        =   9
      Top             =   1830
      Width           =   3630
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   1515
      TabIndex        =   8
      Top             =   1455
      Width           =   1350
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1515
      TabIndex        =   7
      Top             =   1080
      Width           =   1350
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1515
      TabIndex        =   2
      Top             =   705
      Width           =   1350
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   225
      TabIndex        =   0
      Text            =   "Position=12;uid=;Pass=sa;Copyright=B. Cem HANER;Email=info@cemhaner.com"
      Top             =   210
      Width           =   5910
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2865
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   225
      TabIndex        =   6
      Top             =   2250
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   225
      TabIndex        =   5
      Top             =   1875
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   225
      TabIndex        =   4
      Top             =   1485
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   3
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   735
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ASPParser(MyString As String, KeyWord As String, Splitter As String) As String
'Coded by B.Cem HANER info@cemhaner.com
'If you vote me, i can sending very good codes.. Thanx!

Dim KeyWordBul As Integer
Dim DonenSon As Integer

'Attention: This routine perceiving Big and Small character
'Therefore, Your Keyword's spelling is important..
'For example Username=1 is difference username=1

KeyWord = KeyWord + "="
ASPParser = ""
KeyWordBul = InStr(1, MyString, KeyWord)
If KeyWordBul > 0 Then
    DonenSon = InStr(KeyWordBul, MyString, Splitter)
    If DonenSon <= 0 Then
        ASPParser = Mid(MyString, KeyWordBul + Len(KeyWord))
    Else
        ASPParser = Mid(MyString, KeyWordBul + Len(KeyWord), DonenSon - KeyWordBul - Len(KeyWord))
    End If
End If

End Function

Private Sub Command1_Click()

'For example:
'Text2.Text = ASPParser(Text1.Text, "Position", ";")
'                         ¦            ¦         ¦
'                       String      Keyword     Split symbol

Text2.Text = ASPParser(Text1.Text, "Position", ";")
Text3.Text = ASPParser(Text1.Text, "uid", ";")
Text4.Text = ASPParser(Text1.Text, "Pass", ";")
Text5.Text = ASPParser(Text1.Text, "Copyright", ";")
Text6.Text = ASPParser(Text1.Text, "Email", ";")

End Sub

Private Sub Form_Load()
Label2.Caption = "You can use this function;" & vbLf
Label2.Caption = Label2.Caption & "for ASP Data process, for your code process" & vbLf
Label2.Caption = Label2.Caption & "for HTML/XML and other data, needed command line paramter process"
End Sub
