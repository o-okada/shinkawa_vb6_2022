VERSION 5.00
Begin VB.Form �\�񕶃e�X�g���M 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�\�񕶃e�X�g�o�^"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.TextBox Text2 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   4
      Left            =   4305
      TabIndex        =   28
      Text            =   "Text2"
      Top             =   4020
      Width           =   570
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   3
      Left            =   3420
      TabIndex        =   26
      Text            =   "Text2"
      Top             =   4005
      Width           =   570
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   2
      Left            =   2490
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   4005
      Width           =   570
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   1
      Left            =   1590
      TabIndex        =   24
      Text            =   "Text2"
      Top             =   3990
      Width           =   570
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   0
      Left            =   345
      TabIndex        =   19
      Text            =   "Text2"
      Top             =   3990
      Width           =   900
   End
   Begin VB.OptionButton Option1 
      Caption         =   "���ӕ��������"
      Height          =   330
      Index           =   4
      Left            =   1545
      TabIndex        =   17
      Top             =   2040
      Width           =   2685
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   4
      Left            =   4260
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3075
      Width           =   570
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   3
      Left            =   3390
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3075
      Width           =   570
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   2
      Left            =   2490
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3075
      Width           =   570
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   1
      Left            =   1575
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3075
      Width           =   570
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   330
      Index           =   0
      Left            =   345
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3075
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "���ӕ�؊�������"
      Height          =   330
      Index           =   3
      Left            =   1545
      TabIndex        =   5
      Top             =   1596
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      Caption         =   "�^����񕶈�"
      Height          =   330
      Index           =   2
      Left            =   1545
      TabIndex        =   4
      Top             =   1154
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      Caption         =   "�x�񕶈�"
      Height          =   330
      Index           =   1
      Left            =   1545
      TabIndex        =   3
      Top             =   712
      Width           =   2685
   End
   Begin VB.OptionButton Option1 
      Caption         =   "���ӕ񕶈�"
      Height          =   330
      Index           =   0
      Left            =   1545
      TabIndex        =   2
      Top             =   270
      Value           =   -1  'True
      Width           =   2685
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�o�@�^"
      Height          =   495
      Index           =   1
      Left            =   2985
      TabIndex        =   1
      Top             =   4620
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���@�~"
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   4620
      Width           =   1710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   225
      Index           =   9
      Left            =   4035
      TabIndex        =   27
      Top             =   4095
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�N"
      Height          =   225
      Index           =   8
      Left            =   1275
      TabIndex        =   23
      Top             =   4095
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   225
      Index           =   7
      Left            =   2190
      TabIndex        =   22
      Top             =   4080
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   225
      Index           =   6
      Left            =   3090
      TabIndex        =   21
      Top             =   4080
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   225
      Index           =   5
      Left            =   4890
      TabIndex        =   20
      Top             =   4080
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���\����"
      Height          =   225
      Index           =   1
      Left            =   450
      TabIndex        =   18
      Top             =   3705
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   225
      Index           =   4
      Left            =   4905
      TabIndex        =   16
      Top             =   3150
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   225
      Index           =   3
      Left            =   4005
      TabIndex        =   14
      Top             =   3150
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   225
      Index           =   2
      Left            =   3105
      TabIndex        =   12
      Top             =   3150
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   225
      Index           =   1
      Left            =   2205
      TabIndex        =   10
      Top             =   3150
      Width           =   225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "�N"
      Height          =   225
      Index           =   0
      Left            =   1290
      TabIndex        =   8
      Top             =   3165
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���ʎ���"
      Height          =   225
      Index           =   0
      Left            =   540
      TabIndex        =   6
      Top             =   2730
      Width           =   900
   End
End
Attribute VB_Name = "�\�񕶃e�X�g���M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim file_name As String

'Private Sub Command1_Click(Index As Integer)
'    Dim nf   As Integer
'    Dim buf  As String
'    Dim ic   As Boolean
'    Dim st   As String
'    Dim d1   As Date
'    Dim d2   As Date
'    Select Case Index
'        Case 0 '���~
'            Unload Me
'            OracleDB.Timer1.Enabled = True
'        Case 1 '�o�^
'            st = Trim(Text1(0)) & "/" & Trim(Text1(1)) & "/" & Trim(Text1(2)) & " " & _
'                 Trim(Text1(3)) & ":" & Trim(Text1(4))
'            If Not IsDate(st) Then
'                MsgBox "�ݒ肳�ꂽ���ʓ����͓��t�f�[�^�ł͗L��܂���B"
'                Exit Sub
'            End If
'            d1 = CDate(st)
'            st = Trim(Text2(0)) & "/" & Trim(Text2(1)) & "/" & Trim(Text2(2)) & " " & _
'                 Trim(Text2(3)) & ":" & Trim(Text2(4))
'            If Not IsDate(st) Then
'                MsgBox "�ݒ肳�ꂽ���\�����͓��t�f�[�^�ł͗L��܂���B"
'                Exit Sub
'            End If
'            d2 = CDate(st)
'            nf = FreeFile
'            Open App.Path & "\data\" & file_name For Input As #nf
'            C1 = Format(Now, "yyyy/mm/dd hh:nn")
'            Line Input #nf, buf          'FORCAST_KIND        �^���x�񔭕\
'            C2 = Trim(Mid(buf, 21, 20))
'            Line Input #nf, buf          'FORCAST_KIND_CODE   20
'            C3 = Trim(Mid(buf, 21, 20))
'            C4 = Format(d1, "yyyy/mm/dd hh:nn")  'ESTIMATE_TIME
'            C5 = Format(d2, "yyyy/mm/dd hh:nn")  'ANNOUNCE_TIME
'            B1 = ""
'            Do
'                Line Input #nf, buf
'                If Mid(buf, 1, 1) = "*" Then Exit Do
'                B1 = B1 & buf & vbLf
'            Loop
'            B2 = ""
'            Do
'                Line Input #nf, buf
'                If Mid(buf, 1, 1) = "*" Then Exit Do
'                B2 = B2 & buf & vbLf
'            Loop
'            Close #nf
'            ORA_YOHOUBUNAN ic
'    End Select
'    OracleDB.Timer1.Enabled = True
'    MsgBox "�\�񕶂�o�^���܂���"
'    Unload Me
'End Sub

'Private Sub Form_Load()
'    Me.Left = (Screen.Width - Me.Width) * 0.5
'    Me.Top = (Screen.Height - Me.Height) * 0.3
'    Me.Text1(0) = Year(Now)
'    Me.Text1(1) = Month(Now)
'    Me.Text1(2) = Day(Now)
'    Me.Text1(3) = Hour(Now)
'    Me.Text1(4) = 0
'    Me.Text2(0) = Year(Now)
'    Me.Text2(1) = Month(Now)
'    Me.Text2(2) = Day(Now)
'    Me.Text2(3) = Hour(Now)
'    Me.Text2(4) = 0
'    file_name = "���K�p�\�񕶑�P��.dat"
'    Option1(0).Value = True
'End Sub

'Private Sub Option1_Click(Index As Integer)
'    Select Case Index
'        Case 0
'        file_name = "���K�p�\�񕶑�P��.dat"
'        Case 1
'        file_name = "���K�p�\�񕶑�Q��.dat"
'        Case 2
'        file_name = "���K�p�\�񕶑�R��.dat"
'        Case 3
'        file_name = "���K�p�\�񕶑�S��.dat"
'        Case 3
'        file_name = "���K�p�\�񕶑�T��.dat"
'    End Select
'End Sub


