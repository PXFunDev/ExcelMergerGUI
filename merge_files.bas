Attribute VB_Name = "merge_files"
Option Explicit

' -----------------------------------------------------------
' �v���O������: merge_files.bas
' �쐬��: ���Ȃ��̖��O
' �쐬��: 2025�N4��9��
' �o�[�W����: 1.0
' ����: ���̃v���O�����́A�Z�Z�����s���邽�߂̂��̂ł��B
' �g�p���@: ���s���@�̐����������ɋL��
' -----------------------------------------------------------
' �ύX����:
' ���t        �o�[�W����    �ύX���e
' ----------  ----------  -----------------------------------
' 2025/04/09  1.0         ���ō쐬
'
' -----------------------------------------------------------

' -----------------------------------------------------------
' ## �v����`
' - �]�L����Excel�t�@�C�����蓮�I������
' - �]�L���Excel�t�@�C���́A�}�N�������s����Ă���u�b�N�Ƃ���
' - �Ώۂ̃V�[�g�́u�I���T�C�g�v�u�Z���h�o�b�N�v�uN�p�b�P�[�W�v��3��
' - �]�L�̊�́AB��̓o�^�ԍ��Ƃ���
' - �]�L���̂S�s�ڈȍ~���Q��
' - �]�L����B��ڍs���Q��
' - �]�L���̃I���T�C�g�FAR��A�Z���h�o�b�N�FAP��AN�p�b�P�[�W�FAP��܂ł�]�L
' - �]�L���̃f�[�^�����
' - �}�N�����s�̃��O���L�^����
' -----------------------------------------------------------
'************************************************************
' �^�C�g����
'************************************************************
Public Sub merge_files()
  
  On Error GoTo ErrorHandler

    ' �}�N�������s���邩�ǂ����m�F
    Dim response As VbMsgBoxResult
    response = MsgBox("�}�N�������s���܂����H", vbYesNo + vbQuestion, "�m�F")
    If response = vbYes Then

        '************************************************************
        ' ���O����
        '************************************************************
        ' �����J�n���Ԃ��L�^
        Dim T As Double
        T = Timer
        ' �����v�Z����ʍX�V��~
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False

        '************************************************************
        ' ���C���̏���
        '************************************************************
        ' �]�L����Excel�t�@�C��
        Dim SrcWb As Workbook
        Dim SrcWs_Onsite As Worksheet
        Dim SrcWs_Sendback As Worksheet
        Dim SrcWs_NPackage As Worksheet

        ' �C�ӂ�Excel�t�@�C�����J��(�蓮�I��)
        Set SrcWb = Application.Workbooks.Open(Application.GetOpenFilename("Excel Files (*.xls; *.xlsx), *.xls; *.xlsx", , "�Ώۃt�@�C����I�����Ă�������"))
        If SrcWb Is Nothing Then
            MsgBox "�t�@�C�����I������Ă��܂���B�����𒆎~���܂��B", vbExclamation
            Exit Sub
        End If

        ' �Ώۃt�@�C���ɃI���T�C�g�A�Z���h�o�b�N�AN�p�b�P�[�W�����݂��邩�m�F
        Set SrcWs_Onsite = SrcWb.Worksheets("�I���T�C�g")
        Set SrcWs_Sendback = SrcWb.Worksheets("�Z���h�o�b�N")
        Set SrcWs_NPackage = SrcWb.Worksheets("N�p�b�P�[�W")

        ' �]�L���̃t�B���^�[����
        SrcWs_Onsite.AutoFilterMode = False
        SrcWs_Sendback.AutoFilterMode = False
        SrcWs_NPackage.AutoFilterMode = False

        ' �]�L���Excel�t�@�C���i�}�N�������s����Ă���u�b�N�j
        Dim dstWb           As Workbook
        Dim dstWs_Onsite    As Worksheet
        Dim dstWs_Sendback  As Worksheet
        Dim dstWs_NPackage  As Worksheet

        Set dstWb = ThisWorkbook ' �}�N�������s����Ă���u�b�N
        Set dstWs_Onsite = dstWb.Worksheets("�I���T�C�g") ' �]�L��̃I���T�C�g�V�[�g
        Set dstWs_Sendback = dstWb.Worksheets("�Z���h�o�b�N") ' �]�L��̃Z���h�o�b�N�V�[�g
        Set dstWs_NPackage = dstWb.Worksheets("N�p�b�P�[�W") ' �]�L���N�p�b�P�[�W�V�[�g

        ' �Ώۃt�@�C������I���T�C�g�̃f�[�^��]�L
        ' �I���T�C�g�V�[�g�����݂��Ȃ��ꍇ�́A�G���[���b�Z�[�W��\��
        If SrcWs_Onsite Is Nothing Then
            MsgBox "�Ώۃt�@�C���ɃI���T�C�g�V�[�g�����݂��܂���B", vbExclamation
            SrcWb.Close False
            Exit Sub
        Else
            ' �I���T�C�g�V�[�g�����݂���ꍇ�́A�]�L�����s
            Call CheckDataExists(SrcWs_Onsite, dstWs_Onsite)
        End If

        ' �Ώۃt�@�C������Z���h�o�b�N�̃f�[�^��]�L
        ' �Z���h�o�b�N�V�[�g�����݂��Ȃ��ꍇ�́A�G���[���b�Z�[�W��\��
        If SrcWs_Sendback Is Nothing Then
            MsgBox "�Ώۃt�@�C���ɃZ���h�o�b�N�V�[�g�����݂��܂���B", vbExclamation
            SrcWb.Close False
            Exit Sub
        Else
            ' �Z���h�o�b�N�V�[�g�����݂���ꍇ�́A�]�L�����s
            Call CheckDataExists(SrcWs_Sendback, dstWs_Sendback)
        End If

        ' �Ώۃt�@�C������N�p�b�P�[�W�̃f�[�^��]�L
        ' N�p�b�P�[�W�V�[�g�����݂��Ȃ��ꍇ�́A�G���[���b�Z�[�W��\��
        If SrcWs_NPackage Is Nothing Then
            MsgBox "�Ώۃt�@�C����N�p�b�P�[�W�V�[�g�����݂��܂���B", vbExclamation
            SrcWb.Close False
            Exit Sub
        Else
            ' N�p�b�P�[�W�V�[�g�����݂���ꍇ�́A�]�L�����s
            Call CheckDataExists(SrcWs_NPackage, dstWs_NPackage)
        End If

        ' �]�L���̃f�[�^�����(�Z�[�u���Ȃ�)
        Workbooks(SrcWb.Name).Close False

        '************************************************************
        ' �c���
        '************************************************************
        ' �����v�Z����ʍX�V�ĊJ
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        ' ���O: �}�N�����s�̐������L�^
        Call LogMacroExecution("merge_files", "����")
        ' �����������b�Z�[�W
        MsgBox "�}�N�������s���܂����B" & vbCrLf & "��������: " & Format(Timer - T, "0.00") & " �b"
    Else
        ' ���O: �}�N�����s�̃L�����Z�����L�^
        Call LogMacroExecution("merge_files", "�L�����Z��")
        ' �L�����Z�����b�Z�[�W
        MsgBox "�}�N���̎��s���L�����Z�����܂����B"

    End If
    Exit Sub

'************************************************************
' �G���[�n���h�����O
'************************************************************
ErrorHandler:
    ' �]�L���̃f�[�^�����(�Z�[�u���Ȃ�)
    Workbooks(SrcWb.Name).Close False
    ' ���O�F�}�N�����s���A�G���[���b�Z�[�W�����O�ɋL�^
    Call LogMacroExecution("merge_files", "���s - " & Err.Description)
    ' �G���[���b�Z�[�W��\��
    MsgBox "�G���[���������܂����B�Ǘ��҂ɘA�����������B" & vbCrLf & "�G���[���e: " & Err.Description, vbCritical
    ' �����v�Z����ʍX�V�ĊJ
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

' ************************************************************
' �]�L��ɓ]�L���̃f�[�^���Ȃ������؂���T�u���[�`��
Sub CheckDataExists(srcWs As Worksheet, dstWs As Worksheet)
    ' �ϐ��錾���W��
    Dim i As Long
    Dim key As Variant
    Dim srcRegNo As String
    Dim dstRegNo As String
    Dim lastRow As Long
    Dim dstLastRow As Long
    Dim srcDict As Object
    Dim DstDict As Object
    Set srcDict = CreateObject("Scripting.Dictionary")
    Set DstDict = CreateObject("Scripting.Dictionary")

    ' �t�B���^�[������
    srcWs.AutoFilterMode = False
    dstWs.AutoFilterMode = False

    ' �]�L���̓o�^�ԍ����W�v
    lastRow = srcWs.Cells(srcWs.Rows.Count, 2).End(xlUp).Row
    For i = 4 To lastRow
        srcRegNo = srcWs.Cells(i, 2).Value
        If srcRegNo <> "" Then
            If Not srcDict.Exists(srcRegNo) Then
                srcDict.Add srcRegNo, 1
            Else
                srcDict(srcRegNo) = srcDict(srcRegNo) + 1
            End If
        End If
    Next i

    ' �]�L��̓o�^�ԍ����W�v
    dstLastRow = dstWs.Cells(dstWs.Rows.Count, 2).End(xlUp).Row
    For i = 4 To dstLastRow
        dstRegNo = dstWs.Cells(i, 2).Value
        If dstRegNo <> "" Then
            If Not DstDict.Exists(dstRegNo) Then
                DstDict.Add dstRegNo, 1
            Else
                DstDict(dstRegNo) = DstDict(dstRegNo) + 1
            End If
        End If
    Next i

    ' �o�^�ԍ����ƂɌ���r
    Dim needUpdate As Boolean: needUpdate = False
    Dim warnMsg As String: warnMsg = ""
    For Each key In srcDict.Keys
        If Not DstDict.Exists(key) Then
            needUpdate = True
            warnMsg = "�]�L��ɓo�^�ԍ� [" & key & "] ������܂���B"
            Exit For
        ElseIf srcDict(key) <> DstDict(key) Then
            needUpdate = True
            warnMsg = "�o�^�ԍ� [" & key & "] �̌�����v���܂���B" & vbCrLf & _
                        "�]�L��: " & srcDict(key) & "��, �]�L��: " & DstDict(key) & "��"
            Exit For
        End If
    Next key
    For Each key In DstDict.Keys
        If Not srcDict.Exists(key) Then
            needUpdate = True
            warnMsg = "�]�L��ɗ]���ȓo�^�ԍ� [" & key & "] ������܂��B"
            Exit For
        End If
    Next key

    If needUpdate Then
        MsgBox "�]�L���Ɠ]�L��Ńf�[�^�̉ߕs��������܂��B" & vbCrLf & _
                warnMsg & vbCrLf & _
                "�蓮�œ]�L��̃f�[�^���C�����Ă��������B", vbExclamation
        Exit Sub
    End If
    ' �������ׂĈ�v�F�������Ȃ�
    Exit Sub


    ' �]�L���̊Y���s�̂ݓ]�L
    Call CopyData(srcWs, dstWs, 4, lastRow - 3)
End Sub

' ************************************************************
' �f�[�^��]�L����T�u���[�`��

Private Sub CopyData(srcWs As Worksheet, _
            dstWs As Worksheet, _
            targetRow As Long, _
            copyRowCount As Long _
            )

    Dim lastRow As Long
    Dim i As Long
    Dim col As Long

    ' �]�L��̃V�[�g�̍ŏI�s���擾�iB���j
    lastRow = dstWs.Cells(dstWs.Rows.Count, 2).End(xlUp).Row + 1

    ' �v����`�Ɋ�Â��AB��ȍ~��]�L
    Dim colStart As Long, colEnd As Long
    ' �V�[�g���œ]�L�͈͂�؂�ւ�
    Select Case dstWs.Name
        Case "�I���T�C�g"
            colStart = 2 ' B��
            colEnd = 44  ' AR��
        Case "�Z���h�o�b�N"
            colStart = 2 ' B��
            colEnd = 42  ' AP��
        Case "N�p�b�P�[�W"
            colStart = 2 ' B��
            colEnd = 42  ' AP��
    End Select

    For i = targetRow To targetRow + copyRowCount - 1
        For col = colStart To colEnd
            dstWs.Cells(lastRow, col).Value = srcWs.Cells(i, col).Value
        Next col
        lastRow = lastRow + 1
    Next i
End Sub



