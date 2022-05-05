Attribute VB_Name = "MDB_HRdata"
Option Explicit
Option Base 1
Public wH(6, 25)   As Single
Public DH_Tide     As Single

'********************************************************************
'�y�T�u���[�`�����zData_IN()
'�y�����T�v�z
' ����.mdb���琅�ʊϑ��l��ǂݍ��ށB
'�y�����ڍׁz
' (1)����.mdb����1���Ԗ��̐��ʊϑ��l��ǂݍ��ށB
' (2)���ʊϑ��l��-99���܂܂��ꍇ�A��U����B
'�y�C�������z
' 2004/03/24 �����̕�U�͌����������Ƃ����B
'�y�R�����g�z
'��ds�́A�֐��̈�����2021/05/25 00:40���ł���B
'��de�́A�֐��̈�����2021/05/25 17:40���ł���B
'��irc�́A�֐��̖߂�l��True�AFalse�̂����ꂩ�ł���B
'�����̊֐���10������2��Ă΂��悤�ł���B
'��New_RSHINK.exe�ANewNSKG2.exe����1���Ԗ���25���ԕ��A25�̃f�[�^��K�v�Ƃ���悤�ł���B
'********************************************************************
Sub Data_IN(ds As Date, de As Date, irc As Boolean)
    Dim i      As Long
    Dim j      As Integer
    Dim k      As Integer
    Dim m      As Integer
    Dim b      As String
    Dim du     As Date
    Dim dw     As Date
    Dim dur    As Date
    Dim dwr    As Date
    Dim ConR   As New ADODB.Recordset
    Dim a
    Dim SQL    As String
    Dim mi     As String '0,10,20,30,40,50�̉��ꂩ���Z�b�g�����B
    Dim C0     As Single '�V�����ʁi�������j
    Dim C1     As Single '�V�����ʁi1���Ԍ�j
    Dim C2     As Single '�V�����ʁi2���Ԍ�j
    Dim C3     As Single '�V�����ʁi3���Ԍ�j
    Dim ch     As Boolean
    Dim uh     As Boolean
    Dim hw(4)  As Single
    Dim er     As Boolean

    If Err <> 0 Then
        MsgBox "����.MDB�ɃA�N�Z�X�ł��܂���A����.MDB�̗L�����m�F���Ă��������B" & vbCrLf & _
               "�v�Z�ł��܂���̂Ńv���u�����͏I�����܂��B", vbExclamation
        End
    End If

    '****************************************************************
    '�����ϐ��Z�b�g�����i010�j
    '��mi�́A40���ł���B
    '****************************************************************
    mi = Fix(Minute(de) / 10) * 10

    '****************************************************************
    'DB�A�N�Z�X�����i110�j
    '(1)����.mdb����1���Ԗ��̐��ʊϑ��l���擾����B
    '��ds�́A�֐��̈�����2021/05/25 00:40���ł���B
    '��de�́A�֐��̈�����2021/05/25 17:40���ł���B
    '(2)�擾�������ʊϑ��l���O���[�o���ϐ���HO�Ɋi�[����B
    'HO(1,1) ���ʊϑ��l=������O���ʊϑ��l
    'HO(2,1) �􉁉z���ʊϑ��l=0
    'HO(3,1) ���V��F���ʊϑ��l
    'HO(4,1) �厡���ʊϑ��l
    'HO(5,1) �����O���ʊϑ��l
    'HO(6,1) �v�n�쐅�ʊϑ��l
    'HO(7,1) �t�����ʊϑ��l
    '����1�����́A�ϑ����̃C���f�b�N�X�ł���B
    '����2�����́A10:40,11:40,12:40,,,,���̎��n��f�[�^�̃C���f�b�N�X�ł���B
    '����2�����́A����ȏꍇ�A25�����ł���B25���̏ꍇ�A25���ڂ��\���Ώۂ̌������ł���B
    '****************************************************************
    '���ʎ擾
    
    Dim localHOTime(500) As Date
    
    SQL = "select * from ���� where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' and Minute = " & mi & " order by Time"
    Short_Break 4
    ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly
    i = 0
    Do Until ConR.EOF
        dw = ConR.Fields("Time").Value
        If i = 0 Then
            du = dw
        End If
        i = DateDiff("h", du, dw) + 1
        HO(1, i) = ConR.Fields("Tide").Value       'Tide ������O����
        HO(2, i) = 0#                              '�􉁉z����
        HO(3, i) = ConR.Fields("���V��F").Value   '���V��F
        HO(4, i) = ConR.Fields("�厡").Value       '�厡
        HO(5, i) = ConR.Fields("�����O").Value   '����O
        HO(6, i) = ConR.Fields("�v�n��").Value     '�v�n��
        HO(7, i) = ConR.Fields("�t��").Value       '�t��
        localHOTime(i) = ConR.Fields("Time").Value
        ConR.MoveNext
    Loop
    ConR.Close

    '****************************************************************
    '****************************************************************
    '****************************************************************
    '���݂��疢�������̃f�[�^���������i210�j
    '�����̒l���v�Z���āA�����ϐ���DH_Tide�ɃZ�b�g����B
    '(1)�\���v�Z�̌������̒��ʊϑ��l��-99�̏ꍇ�ŁA10���O�̋C�ے������̓�����O���ʂ������̏ꍇ�ADH_Tide=0
    '(2)�\���v�Z�̌������̒��ʊϑ��l��-99�̏ꍇ�ŁA10���O�̋C�ے������̓�����O���ʌv�Z�l������ꍇ�ADH_Tide=10���O�̓�����O���ʌv�Z�l
    '(3)�\���v�Z�̌������̒��ʊϑ��l��-99�ȊO�̏ꍇ�A�A�ADH_Tide=0
    '��(3)��DH_Tide�͌��ǂ̂Ƃ���g�p���Ȃ��BDH_Tide���g�p��������A�܂�HO(1,Now_Step)<-50#�̂��߁B
    '�y�C�������z
    ' 2022/03/08 �l������̃o�O�����邽�߁A�R�����g�A�E�g�����B
    '****************************************************************
    '****************************************************************
    '****************************************************************
    '''If HO(1, Now_Step) < -50# Then
    '''    Tide_Special
    '''    ORA_Message_Out "���ʃf�[�^��M", "������O���ʃf�[�^���������܂����B�V�����ʂɒ��O�̎����l�Ƃ̍����������āA�����E�\���l�Ƃ��܂��B", 1
    '''Else
    '''    DH_Tide = 0#
    '''End If
    
    Dim localIndex              As Integer
    Dim localSQL                As String
    Dim localObsTime(3)         As Date
    Dim localObsTide(3)         As Single
    Dim localTenmonTide(3)      As Single
    Dim localObsMinusTenmonTide As Single
    Dim localTenmonTide1        As Single
    Dim localTenmonTide2        As Single
    Dim localTenmonTide3        As Single
    
    localSQL = "SELECT TOP 3 TIME,TIDE FROM ���� WHERE TIME < '" & Format(de, "yyyy/mm/dd hh:nn") & _
               "' AND TIDE <> -99 ORDER BY TIME DESC"
    Short_Break 1
    ConR.Open localSQL, Con_����, adOpenKeyset, adLockReadOnly
    localIndex = 1
    Do Until ConR.EOF
        localObsTime(localIndex) = ConR.Fields("TIME").Value
        localObsTide(localIndex) = ConR.Fields("TIDE").Value '���ʊϑ��l=������O���ʊϑ��l
        localIndex = localIndex + 1
        ConR.MoveNext
    Loop
    ConR.Close
    
    Cal_Tide localObsTime(1), localTenmonTide(1), localTenmonTide1, localTenmonTide2, localTenmonTide3
    Cal_Tide localObsTime(2), localTenmonTide(2), localTenmonTide1, localTenmonTide2, localTenmonTide3
    Cal_Tide localObsTime(3), localTenmonTide(3), localTenmonTide1, localTenmonTide2, localTenmonTide3
    localObsMinusTenmonTide = ((localObsTide(1) - localTenmonTide(1)) + (localObsTide(2) - localTenmonTide(2)) + (localObsTide(3) - localTenmonTide(3))) / 3#
    DH_Tide = ((localObsTide(1) - localTenmonTide(1)) + (localObsTide(2) - localTenmonTide(2)) + (localObsTide(3) - localTenmonTide(3))) / 3#
    LOG_Out "de=" & de & "localObsMinusTenmonTide=" & localObsMinusTenmonTide
    
    '������.mdb�̋C�ے������e�[�u���̓�����O���ʃJ�����̑�1�v�f��DH_Tide���o�͂���邽�߁A����݊����̂��߂�DH_Tide���Z�b�g����B

    '****************************************************************
    '���݂��疢�������̃f�[�^���������i220�j
    '�\���v�Z�̌������̓V�����ʁi�������A1���Ԍ�A2���Ԍ�A3���Ԍ�j���v�Z����B
    '��dw�́A�\���v�Z�̌������̒l�ł���B
    '��C0�́A�V�����ʂ̌������̒l�ł���B
    '��C1�́A�V�����ʂ�1���Ԍ�̒l�ł���B
    '��C2�́A�V�����ʂ�2���Ԍ�̒l�ł���B
    '��C3�́A�V�����ʂ�3���Ԍ�̒l�ł���B
    '****************************************************************
    '''�\�����ʗՎ�
    '''TidalY dw, C0, C1, C2, C3    '�C�ے����ʕ\����V�����ʂ���}����
    Cal_Tide dw, C0, C1, C2, C3     '60��������V�����ʂ��v�Z����
    
    '****************************************************************
    '****************************************************************
    '****************************************************************
    '���݂��疢�������̃f�[�^���������i230�j
    '(1)�\���v�Z�̌������̒��ʊϑ��l��-99�̏ꍇ�ŁA10���O�̋C�ے������̓�����O���ʂ������̏ꍇ�ADH_Tide=0�ŁA�V�����ʂ�1���Ԍ�A2���Ԍ�A3���Ԍ�̒��ʗ\���l�Ƃ���B
    '(2)�\���v�Z�̌������̒��ʊϑ��l��-99�̏ꍇ�ŁA10���O�̋C�ے������̓�����O���ʂ�����ꍇ�ADH_Tide=10���O�̓�����O���ʌv�Z�l�ŁA1���Ԍ�A2���Ԍ�A3���Ԍ�̒��ʗ\���l�Ƃ���B
    '(3)�\���v�Z�̌������̒��ʊϑ��l��-99�ȊO�̏ꍇ�A���ʊϑ��l�̌������̒l�ɓV�����ʂ�1���ԕ��A2���ԕ��A3���ԕ��̕ϓ��l�����Z���āA1���Ԍ�A2���Ԍ�A3���Ԍ�̒��ʗ\���l�Ƃ���B
    '��(1)(2)�Ƃ��l���R��̃o�O�ŏ�Q����������B
    '****************************************************************
    '****************************************************************
    '****************************************************************
    If HO(1, Now_Step) < -50# Then
        '''HO(1, Now_Step) = C0 + DH_Tide
        '''HO(1, Now_Step + 1) = C1 + DH_Tide
        '''HO(1, Now_Step + 2) = C2 + DH_Tide
        '''HO(1, Now_Step + 3) = C3 + DH_Tide
        HO(1, Now_Step) = C0 + localObsMinusTenmonTide
        HO(1, Now_Step + 1) = C1 + localObsMinusTenmonTide
        HO(1, Now_Step + 2) = C2 + localObsMinusTenmonTide
        HO(1, Now_Step + 3) = C3 + localObsMinusTenmonTide
    Else
        '''DH_Tide = HO(1, Now_Step) - C0
        '''HO(1, Now_Step) = HO(1, Now_Step) + C0 - C0
        HO(1, Now_Step + 1) = HO(1, Now_Step) + C1 - C0
        HO(1, Now_Step + 2) = HO(1, Now_Step) + C2 - C0
        HO(1, Now_Step + 3) = HO(1, Now_Step) + C3 - C0
    End If
    
    '****************************************************************
    '���݂��疢�������̃f�[�^���������i240�j
    '�����L240�̏������R�����g�A�E�g���āA�����ȏ����ƂȂ�悤�ɏ�L230�̏������C�������B
    '****************************************************************
    '''HO(1, Now_Step + 1) = C1 + DH_Tide
    '''HO(1, Now_Step + 2) = C2 + DH_Tide
    '''HO(1, Now_Step + 3) = C3 + DH_Tide

    '****************************************************************
    '�������菈���i310�j
    '����.mdb����擾�������ʊϑ��l�̌�����0���̏ꍇ�A�֐��𔲂���B
    '****************************************************************
    If i = 0 Then
        '''MsgBox "���[�J��DB�ɐ��ʃf�[�^������܂���B"
        LOG_Out "���[�J��DB�ɐ��ʃf�[�^������܂���B"
        ds = CDate("1900/01/01 01:00")
        de = CDate("1900/01/01 01:00")
        Exit Sub
    End If

    '****************************************************************
    '�����ϐ��Z�b�g�����i320�j
    '****************************************************************
    '''���ѐ��ʍŏI�f�[�^���t�̗\���f�[�^�����ɍs��
    Set ConR = Nothing
    jsd = du
    js(1) = Year(jsd)
    js(2) = Month(jsd)
    js(3) = Day(jsd)
    js(4) = Hour(jsd)
    js(5) = Minute(jsd)
    js(6) = 0
    jgd = dw
    jg(1) = Year(jgd)
    jg(2) = Month(jgd)
    jg(3) = Day(jgd)
    jg(4) = Hour(jgd)
    jg(5) = Minute(jgd)
    jg(6) = 0
    Now_Step = DateDiff("h", jsd, jgd) + 1
    All_Step = Now_Step + Yosoku_Step

    '****************************************************************
    '���O�o�͏����i330�j
    '�v�Z�Ɏg�p���鐅�ʊϑ��l�̌��������Ȃ��ꍇ�Ƀ��O���o�͂���B
    '��Now_Step��DB�A�N�Z�X�����Ŏ擾����1���Ԗ��̐��ʊϑ��l�̌����ł���B
    '��Now_Step�͗Ⴆ�΁A10:40,11:40,12:40,,,,�̐��ʊϑ��l���擾����B
    '��Now_Step�͐���ȏꍇ�A25�����ł���B
    '''�����������Ȃ��ꍇ�Airc���X�V�����ɁA�����Ŋ֐��𔲂���B
    '****************************************************************
    If Now_Step <= 4 Then
        '''LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��I�I�I"
        '''�C���J�n�@2016/09/23�@O.OKADA�@��������R�����g�A�E�g����B
        '''�C�����R�@�v�Z���������15�����x�x��Ă��邽�߁B
        '''Exit Sub
        '''�C���I���@2016/09/23�@O.OKADA�@�����܂ŃR�����g�A�E�g����B
    End If
    
    '****************************************************************
    '���O�o�͏����i340�j
    '�v�Z�Ɏg�p���鐅�ʊϑ��l�̌��������Ȃ��ꍇ�Ƀ��O���o�͂���B
    '��Now_Step��DB�A�N�Z�X�����Ŏ擾����1���Ԗ��̐��ʊϑ��l�̌����ł���B
    '��Now_Step�͗Ⴆ�΁A10:40,11:40,12:40,,,,�̐��ʊϑ��l���擾����B
    '��Now_Step�͐���ȏꍇ�A25�����ł���B
    '''�����������Ȃ��ꍇ�Airc���X�V�����ɁA�����Ŋ֐��𔲂���B
    '****************************************************************
    If ds = de Or All_Step < 3 Then
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        LOG_Out "IN  Data_IN  �v�Z�Ɏg�p���鐅�ʃf�[�^�̃X�e�b�v�������Ȃ��̂Ōv�Z���~���Ȃ��H�H�H"
        '''�C���J�n�@2016/09/23�@O.OKADA�@��������R�����g�A�E�g����B
        '''�C�����R�@�v�Z���������15�����x�x��Ă��邽�߁B
        '''Exit Sub
        '''�C���I���@2016/09/23�@O.OKADA�@�����܂ŃR�����g�A�E�g����B
    End If

    '****************************************************************
    '****************************************************************
    '****************************************************************
    '�ߋ������̃f�[�^���������i410�j
    '���܂�A��U�����ł���B
    'HO(1,i) ���ʊϑ��l=������O���ʊϑ��l
    '����1�����́A�ϑ����̃C���f�b�N�X�ł���B
    '����2�����́A10:40,11:40,12:40,,,,���̎��n��f�[�^�̃C���f�b�N�X�ł���B
    '����2�����́A����ȏꍇ�A25�����ł���B25���̏ꍇ�A25���ڂ��\���Ώۂ̌������ł���B
    '�����ʊϑ��l=������O���ʊϑ��l��-99�̏ꍇ�A
    '����2�������A�ŏ��̒l�̏ꍇ�A���ʊϑ��l�̕��ϒl�Ƃ��Ă���1.5m�ɃZ�b�g���邪���ł���B
    '����2�������A�ŏ��̒l�ȊO�̏ꍇ�A���ʊϑ��l��1�O�̒l���R�s�[���Ă��邪���ł���B
    '�s��́A2021/05/25 18:40�ɔ�������B
    '�����́A10���O�̋C�ے������̓�����O���ʊϑ��l�������Ō����ɂȂ����ꍇ�A
    'HO(1,i)=HO(1,i-1)�Ƃ��Ă��邽�߁A�f�[�^������̒��ʂ�����-99�����O�̒l�ƂȂ邽�߂ł���B
    '�Ⴆ�΁A���ʂ̎��n��f�[�^��HO(1,i)��1.5,0.07,0.07,,,,,,,0.07,1.5���ƂȂ�B
    '3���ԑO�̒l�ƌ������̐��ʊϑ��l�̍�����p���ăX���C�h�������Ă��邽�߁A
    '���ʂ��{��1.5m�ŗǂ��ɂ��ւ�炸�A1.5+(1.5-0.07)=2.93m���ƂȂ�B
    '****************************************************************
    '****************************************************************
    '****************************************************************
    '''�����[���E����������O���ʂ̕�U
    '''�C��1
    '''For i = 1 To Now_Step
    '''    If HO(1, i) < -50# Then
    '''        j = 1
    '''        Select Case i
    '''            Case 1
    '''                HO(1, 1) = 1.5
    '''            Case Is > 1
    '''                HO(1, i) = HO(1, i - 1)
    '''        End Select
    '''    End If
    '''Next i
    Dim localHOTenmonTide0 As Single
    Dim localHOTenmonTide1 As Single
    Dim localHOTenmonTide2 As Single
    Dim localHOTenmonTide3 As Single
    For i = 1 To Now_Step
        If HO(1, i) < -50# Then
            Cal_Tide localHOTime(i), localHOTenmonTide0, localHOTenmonTide1, localHOTenmonTide2, localHOTenmonTide3
            HO(1, i) = localHOTenmonTide0 + localObsMinusTenmonTide
        End If
    Next i

    '****************************************************************
    '�ߋ������̃f�[�^���������i420�j
    '���܂�A��U�����ł���B
    '���ʊϑ��l=������O���ʊϑ��l����t�����ʊϑ��l�̗\���v�Z�̌������̒l��-99�̏ꍇ�AJ1�̏������s���B
    '��i=1�́A���ʊϑ��l=������O���ʊϑ��l
    '��i=2�́A�􉁐��ʊϑ��l
    '��i=3�́A���V��F�A�A�A�ł���B
    '��j=Now_Step�́A�\���v�Z�̌������ł���B
    '���ߋ������̃f�[�^�����̕K�v���𔻒肵�Ă��邪�A���̌�̏��������ʕs���A�e���s���A�Ӗ��s���ł���B
    '****************************************************************
    '''������U
    er = False
    For i = 1 To 7
        If i <> 2 Then
            ch = False
            '********************************************************
            '''For j = Now_Step - 3 To Now_Step '2004/03/24
            '********************************************************
            For j = Now_Step To Now_Step
                a = HO(i, j)
                If a < -50# Then
                    ch = True
                    GoTo J1
                End If
            Next j
        End If
    Next i
J1:
    If ch Then
        '************************************************************
        '�ߋ������̃f�[�^���������i430�j
        '���܂�A��U�����ł���B
        '��wH()�ɒl���Z�b�g���Ă��邪�A���ʕs���A�e���s���A�Ӗ��s���ł���B
        '************************************************************
        Pre_���ʌ�����U
        
        '************************************************************
        '�ߋ������̃f�[�^���������i440�j
        '���܂�A��U�����ł���B
        '�������ϐ���hw()�ɒl���Z�b�g���Ă��邪�A���ʕs���A�e���s���A�Ӗ��s���ł���B
        '************************************************************
        For i = 1 To 7
            If i <> 2 Then
                uh = True
                '****************************************************
                '''For j = Now_Step - 3 To Now_Step '2004/03/24
                '****************************************************
                For j = Now_Step To Now_Step
                    a = HO(i, j)
                    If a < -50# Then
                        uh = True
                        '********************************************
                        '''For k = Now_Step - 3 To Now_Step '2004/03/24
                        '********************************************
                        For k = Now_Step To Now_Step
                            '****************************************
                            '''m = k - (Now_Step - 3) + 1 '2004/03/24
                            '****************************************
                            m = k - Now_Step + 1
                            hw(m) = HO(i, k)
                        Next k
                        If hw(m) < -50# Then
                            er = True
                            ORA_Message_Out "�e�����[�^���ʎ�M", Name_H(i) & "�́A���ʃf�[�^���������܂����B�^���\���V�X�e���ɂ�錋�ʂ�p���Đ��ʗ\���v�Z���s���܂��B", 1
                        End If
                        Exit For
                    End If
                Next j
            End If
        Next i
    End If
    
    '****************************************************************
    '�G���[�����i510�j
    '�\���v�Z�̌������̒��ʊϑ��l=������O���ʊϑ��l������-99�̏ꍇ�A�����ϐ��̌���-99����ϐ���True���Z�b�g����B
    '****************************************************************
    If HO(1, Now_Step) < -50# Then
        er = True
    End If
    
    '****************************************************************
    '�G���[�����i520�j
    '�߂�l��irc��True���Z�b�g����B
    '****************************************************************
    irc = True
    
    '****************************************************************
    '�G���[�����i530�j
    '����͔��菈���A����-99���菈��
    '������͂���`�F�b�N�{�b�N�X���`�F�b�N����Ă��āA�\���v�Z�̌������̒��ʊϑ��l=������O���ʊϑ��l������-99�̏ꍇ�A����͗p�̊֐����Ăяo���B
    '������������ʂ̂ݗL���Ƃ��Ă��邽�߁AIF���̓����̏����͎��s����Ȃ��B
    '******************************************************
    If (AutoDrive.Check6 = vbChecked) And er Then
        '''������U������͂���
        Load Data_Edit
        Unload Data_Edit
    End If
    
    '****************************************************************
    '�G���[�����i540�j
    '��IF���̓����̏����͎��s����Ȃ��B
    '��IF���̓����̏����̓R�����g�A�E�g�ς݂ł���B
    '****************************************************************
    If (AutoDrive.Check6 = vbUnchecked) And er Then
        '''�����Ȃ̂Ōv�Z���X�L�b�v����
        '''irc = False '�����ł��v�Z����悤�ɏC�� 2004/4/26
        '''Exit Sub
    End If

    '''Dim nf As Long
    '''
    '''nf = FreeFile
    '''open app.Path & "\data\���ʃX���C�h��.dat" for output
    '''LOG_Out "IN  Data_IN  ���ʃX���C�h�� CX=" & Format(cx, "###0.000")
    '''MDB_�� jsd, jgd, er
End Sub

'********************************************************************
'�y�T�u���[�`�����zPre_���ʌ�����U()
'�y�����T�v�z
'�y�����ڍׁz
'�y�C�������z
'********************************************************************
Sub Pre_���ʌ�����U()
    Dim ConR        As New ADODB.Recordset
    Dim SQL         As String
    Dim ds          As Date
    Dim de          As Date
    Dim i           As Long
    Dim j           As Long

    '****************************************************************
    '�����ϐ��Z�b�g�����i010�j
    '****************************************************************
    ds = DateAdd("h", -4, jgd)
    de = jgd

    '****************************************************************
    'DB�A�N�Z�X�����i020�j
    '****************************************************************
    SQL = "select * from ���� where Time between '" & Format(ds, "yyyy/mm/dd hh:nn") & "' and '" & _
           Format(de, "yyyy/mm/dd hh:nn") & "' order by Time"
    ConR.Open SQL, Con_����, adOpenKeyset, adLockReadOnly
    j = 1
    Do Until ConR.EOF
        For i = 1 To 6 '6���ʊϑ���
            wH(i, j) = ConR.Fields(i + 1).Value
        Next i
        j = j + 1
        ConR.MoveNext
    Loop
    ConR.Close

End Sub

'********************************************************************
'�y�T�u���[�`�����zTide_Special()
'�y�����T�v�z
'�y�����ڍׁz
'�y�C�������z
' 2022/03/08 ���ʊϑ��l=������O���ʊϑ��l��-99�ŁA
' 10���O�̋C�ے������̓�����O���ʊϑ��l�������ɂȂ����ꍇ�A�\�����ʂ�1.5m���x�}�㏸����s�����������B
' �s��́A2021/05/28 08:30�ɔ�������B
' �����́A10���O�̋C�ے������̓�����O���ʊϑ��l�������ɂȂ����ꍇ�ADH_Tide=0�ƂȂ邽�߂ł���B
' �Ή����邽�߁ASQL����10���O�̐������폜���āA�ŐV�̋C�ے������̓�����O���ʊϑ��l���擾����悤�ɏC�������B
'********************************************************************
Sub Tide_Special()
    Dim SQL    As String
    Dim buf    As String
    Dim dw     As Date
    Dim w

    '****************************************************************
    '���O�o�͏����i010�j
    '****************************************************************
    LOG_Out "IN   Tide_Special"
    
    '****************************************************************
    '�G���[�����i020�j
    '****************************************************************
    On Error GoTo ER1

    '****************************************************************
    '�����ϐ��Z�b�g�����i030�j
    '****************************************************************
    DH_Tide = 0#
    '''MDB_����_Connection
    dw = DateAdd("n", -10, jgd)

    '****************************************************************
    'DB�A�N�Z�X�����i040�j
    '****************************************************************
    If isRAIN = "02" Then
        SQL = "SELECT ������O���� FROM FRICS���� WHERE TIME='" & Format(dw, "yyyy/mm/dd hh:nn") & "'"
    Else
        '************************************************************
        '************************************************************
        '************************************************************
        '''SQL = "SELECT ������O���� FROM �C�ے����� WHERE TIME='" & Format(dw, "yyyy/mm/dd hh:nn") & "'"
        '�y�C�������z
        ' 2022/03/08 ���ʊϑ��l=������O���ʊϑ��l��-99�ŁA
        ' 10���O�̋C�ے������̓�����O���ʊϑ��l�������ɂȂ����ꍇ�A�\�����ʂ�1.5m���x�}�㏸����s�����������B
        ' �s��́A2021/05/28 08:30�ɔ�������B
        ' �����́A10���O�̋C�ے������̓�����O���ʊϑ��l�������ɂȂ����ꍇ�ADH_Tide=0�ƂȂ邽�߂ł���B
        ' �Ή����邽�߁ASQL����10���O�̐������폜���āA�ŐV�̋C�ے������̓�����O���ʊϑ��l���擾����悤�ɏC�������B
        '************************************************************
        '************************************************************
        '************************************************************
        SQL = "SELECT ������O���� FROM �C�ے����� WHERE TIME IN (SELECT MAX(TIME) FROM �C�ے�����)"
    End If
    Rec_����.Open SQL, Con_����, adOpenDynamic, adLockReadOnly
    If Rec_����.EOF Then
        DH_Tide = 0#
    Else
        buf = Rec_����.Fields(0).Value
        w = Split(buf, ",")
        DH_Tide = w(0)
    End If
    Rec_����.Close
    '''MDB_����_Close

    '****************************************************************
    '���O�o�͏����i050�j
    '****************************************************************
    LOG_Out "OUT  Tide_Special DH_Tide=" & Format(DH_Tide, "###0.000")
    Exit Sub

ER1:
    '****************************************************************
    '�G���[�����i060�j
    '****************************************************************
    LOG_Out "OUT  Tide_Special ABend DH_Tide=" & Format(DH_Tide, "###0.000")
    Rec_����.Close
    On Error GoTo 0
End Sub