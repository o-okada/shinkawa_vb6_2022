Attribute VB_Name = "VSPrinter_Primitve"
Option Explicit
Public VSobj    As Object
Public Wbyte    As String
Public Sbyte    As String

Function Cvt_2byte(str As String)
    Dim i As Integer
    For i = 1 To Len(str)
       Cvt_2byte = Cvt_2byte & Mid(Wbyte, InStr(Sbyte, Mid(str, i, 1)), 1)
    Next i
End Function

Sub Int_Cvt_2byte()
    Sbyte = "0123456789/.-+abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ( )"
    Wbyte = "�O�P�Q�R�S�T�U�V�W�X�^�D�|�{����������������������������������������������������" & _
            "�`�a�b�c�d�e�f�g�h�i�j�k�l�m�n�o�p�q�r�s�t�u�v�w�x�y�i�@�j"
End Sub

'----------------------------------------------------------
'x,y,r(mm)��Ίp���ɂ����~��`��
'c      �����F
'w      ����葾��(mm) single
'BrushC �h��Ԃ��F
'BrushS 0=�h��Ԃ��@�@�@�@4=�ΐ�(�E��)
'       1=�����@�@�@�@�@�@�@5=�ΐ�(�E��)
'       2=�������@�@�@�@�@�@6=����(����)
'       3=�������@�@�@�@�@�@7=����(��)
'-----------------------------------------------------------
Sub VS_Circle(x As Single, y As Single, r As Single, _
              c As Integer, w As Single, BrushC As Long, BrushS As Integer)
    VSobj.BrushColor = BrushC
    VSobj.BrushStyle = BrushS
    VS_LineColor c
    VS_LineWidth w
    VSobj.DrawCircle MM(x), MM(y), MM(r)
End Sub

Sub VS_End()
   VSobj.EndDoc
End Sub

Sub VS_Font(F As Integer)
    Dim Fon As String
    Select Case F
        Case 1
            'Ver1.0.0 �C���J�n 2015/08/08 O.OKADA
            'Fon = "�l�r �S�V�b�N"
            Fon = "MS-PGothic"
            'Ver1.0.0 �C���I�� 2015/08/08 O.OKADA
        Case 2
            'Ver1.0.0 �C���J�n 2015/08/08 O.OKADA
            'Fon = "�l�r ����"
            Fon = "MS-PGothic"
            'Ver1.0.0 �C���I�� 2015/08/08 O.OKADA
        Case 3
            Fon = "Times New Roman"
        Case Else
            'Ver1.0.0 �C���J�n 2015/08/08 O.OKADA
            'Fon = "�l�r �S�V�b�N"
            Fon = "MS-PGothic"
            'Ver1.0.0 �C���I�� 2015/08/08 O.OKADA
    End Select
    VSobj.FontName = Fon
End Sub

Sub VS_LineWidth(w As Single)
    'Ver1.0.0 �C���J�n 2015/08/08 O.OKADA
    'VSobj.PenWidth = MM(w)
    VSobj.PenWidth = 10
    'Ver1.0.0 �C���I�� 2015/08/08 O.OKADA
End Sub

Sub VS_NewPage()
    VSobj.Action = paNewPage
End Sub

Sub VS_number(x As Single, y As Single, siz As Single, F As Single, _
             nf As Integer, ic As Integer)
    Dim ww  As String
    Select Case nf
        Case -1
            ww = Format(F, "########0")
        Case 0
            ww = Format(F, "#######0.")
        Case 1
            ww = Format(F, "######0.0")
        Case 2
            ww = Format(F, "#####0.00")
        Case 3
            ww = Format(F, "####0.000")
        Case 4
            ww = Format(F, "###0.0000")
    End Select
    VS_symbol x, y, siz, Cvt_2byte(ww), ic
End Sub

Sub VS_Open()
    VSobj.Preview = True
    VSobj.PaperSize = pprA4
'    VSobj.ExportFormat = vpxPlainHTML
    VSobj.Orientation = orLandscape
    VSobj.StartDoc
    VSobj.MarginTop = 0
    VSobj.MarginBottom = 0
    VSobj.MarginRight = 0
    VSobj.MarginLeft = 0
    VS_Font 2
End Sub

Function MM(F As Single) As Variant
    MM = Format(F, "####0.000mm")
End Function

'----------------------------------------------------------
'X1,Y1,X2,Y2(mm)��Ίp���ɂ�����(��)���`��`��
'LineC  �����F
'LineW  ����葾��(mm) single
'BrushC �h��Ԃ��F
'BrushS 0=�h��Ԃ��@�@�@�@4=�ΐ�(�E��)
'       1=�����@�@�@�@�@�@�@5=�ΐ�(�E��)
'       2=�������@�@�@�@�@�@6=����(����)
'       3=�������@�@�@�@�@�@7=����(��)
'-----------------------------------------------------------
Sub VS_Box(x1 As Single, y1 As Single, x2 As Single, y2 As Single, _
           LineC As Long, LineW As Single, BrushC As Long, BrushS As Integer)
    VSobj.x1 = MM(x1)
    VSobj.y1 = MM(y1)
    VSobj.x2 = MM(x2)
    VSobj.y2 = MM(y2)
    VSobj.BrushColor = BrushC
    VSobj.BrushStyle = BrushS
    VSobj.PenColor = LineC
    VSobj.PenWidth = MM(LineW)
    VSobj.Draw = doRectangle
    VSobj.PenWidth = 0
End Sub

Sub VS_Line(x1 As Single, y1 As Single, x2 As Single, y2 As Single, c As Integer, w As Single)
    VS_LineColor c
    VS_LineWidth w
    VSobj.DrawLine MM(x1), MM(y1), MM(x2), MM(y2)
End Sub

'----------------------------------------------------
'X1(i),Y1(i)(mm)�z����W�̐��`�ʂ��s��
'N ���W��
'C ���F QBColor�̔ԍ�
'W ���� mm
'-----------------------------------------------------
Sub VS_Lines(x1() As Single, y1() As Single, n As Integer, c As Integer, w As Single)
    Dim i As Integer
    Dim p As String
    VS_LineColor c
    VS_LineWidth w
    p = ""
    For i = 1 To n
        p = p & MM(x1(i)) & " " & MM(y1(i)) & " "
    Next i
    VSobj.Polyline = p
End Sub

'-----------------------------------------------
'c=0  ���@�@�@�@8   �D�F
'  1  �@�@�@�@9   ���邢��
'  2  �΁@�@�@�@10  ���邢��
'  3  �V�A���@�@11  ���邢�V�A��
'  4  �ԁ@�@�@�@12  ���邢��
'  5  �}�[���^�@13  ���邢�}�[���^
'  6  ���@�@�@�@14  ���邢��
'  7  ���@�@�@�@15  ���邢��
'-----------------------------------------------
Sub VS_LineColor(c As Integer)
    VSobj.PenColor = QBColor(c)
End Sub

Sub VS_cntlna(n As Integer, al As Single, jp As Integer, XA As Single, ya As Single, XB As Single, _
              yb As Single, irc As Integer, jj As Integer)
      Dim c As Single, S As Single
      Dim dlx As Single, dly As Single, el As Single, XC As Single, YC As Single
      irc = 0
      jj = n
      dlx = XB - XA
      dly = yb - ya
      el = Sqr(dlx * dlx + dly * dly)
      If el = 0# Then
           irc = 1
           Exit Sub
      End If
      If al >= el Then
           If jp = 2 Then VSobj.DrawLine MM(XA), MM(ya), MM(XB), MM(yb)
           al = al - el
           XA = XB
           ya = yb
           irc = 1
           Exit Sub
      Else
           c = dlx / el
           S = dly / el
           XC = al * c + XA
           YC = al * S + ya
           If jp <> 3 Then
               VSobj.DrawLine MM(XA), MM(ya), MM(XC), MM(YC)
           End If
           XA = XC
           ya = YC
      End If
End Sub

Sub VS_cntlnw(x As Single, y As Single, a As Single, b As Single, c As Single, _
              L As Integer, n As Integer)
      Static XA As Single, ya As Single, XB As Single, yb As Single
      Static jj As Integer, d As Single, irc As Integer
      Dim i As Integer
    XB = x
    yb = y
    If n = 1 Then
'        VSobj.x = MM(x)
'        VSobj.y = MM(y)
        XA = x
        ya = y
        jj = 1
        d = a
        Exit Sub
    End If
    If jj = 1 Then GoTo L10
    If jj = 2 Then GoTo L20
    If jj = 3 Then GoTo L30
    If jj = 4 Then GoTo L40
L10:        Call VS_cntlna(1, d, 2, XA, ya, XB, yb, irc, jj)
            If irc <> 0 Then Exit Sub
            d = b
            i = 1
L12:        If i - L > 0 Then GoTo L32
L20:        Call VS_cntlna(2, d, 3, XA, ya, XB, yb, irc, jj)
            If irc <> 0 Then Exit Sub
            d = c
L30:        Call VS_cntlna(3, d, 2, XA, ya, XB, yb, irc, jj)
            If irc <> 0 Then Exit Sub
            d = b
            i = i + 1
            GoTo L12
L32:        d = b
L40:        Call VS_cntlna(4, d, 3, XA, ya, XB, yb, irc, jj)
            If irc <> 0 Then Exit Sub
            d = a
            GoTo L10
End Sub

Sub VS_ShowPage(n As Integer)
    VSobj.PreviewPage = n
End Sub

'--------------------------------------------
'x,y(mm)���W�Ŏw�����ꂽ�Ƃ���ɕ���������
'size �|�C���g
'moji$ �`�敶��
'ic  1     4     7
'    2     5     8
'    3     6     9
'--------------------------------------------
Sub VS_symbol(x As Single, y As Single, size As Single, moji$, ic As Integer)
    Dim xx As Single, yy As Single, Texth As Single, Textw As Single
    VSobj.FontSize = size
    Textw = VSobj.TextWidth(moji$)
    Texth = VSobj.TextHeight(moji$)
    Textw = Len(moji$) * size / 2.9
    Texth = size / 2.9
    Select Case ic
        Case 1, 2, 3
            xx = x
        Case 4, 5, 6
            xx = x - Textw * 0.5
        Case 7, 8, 9
            xx = x - Textw
    End Select
    Select Case ic
        Case 1, 4, 7
            yy = y
        Case 2, 5, 8
            yy = y - Texth * 0.5
        Case 3, 6, 9
            yy = y - Texth
    End Select
    VSobj.CurrentX = MM(xx)
    VSobj.CurrentY = MM(yy)
    VSobj.Text = moji$
End Sub

Sub VS_symbolw(x1, x2, y, size, moji$)
    Dim xx As Single
    Dim yy As Single
    Dim Texth As Single
    Dim Textw As Single
    Dim w As Single, XP As Single, i  As Integer
    'Ver1.0.0 �C���J�n 2015/08/08 O.OKADA
    'VSobj.FontSize = size
    VSobj.FontSize = 12
    'Ver1.0.0 �C���I�� 2015/08/08 O.OKADA
    Texth = VSobj.TextHeight(moji$) / 2.9
    Textw = VSobj.TextWidth(Mid$(moji$, 1, 1)) / 2.9
    w = Abs(x2 - x1)
    xx = (x1 + x2) * 0.5 - w * 0.5
    yy = y - Texth * 0.5
    If Len(moji$) - 1 = 0 Then
        VSobj.x = xx
        VSobj.y = yy
        VSobj.Print moji$
        Exit Sub
    End If
    XP = (w - Textw) / (Len(moji$) - 1)
    For i = 1 To Len(moji$)
        VSobj.x = MM(xx)
        VSobj.y = MM(yy)
        VSobj.Text Mid$(moji$, i, 1)
        xx = xx + XP
    Next i
End Sub
