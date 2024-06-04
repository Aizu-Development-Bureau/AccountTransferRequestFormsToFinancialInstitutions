Attribute VB_Name = "Module1"

' �t�B���^�����O���ꂽ�f�[�^���擾����֐�
' ws: �����Ώۂ̃V�[�g
' filterWork: �������[�h
Function GetFilteredData(ws As Worksheet, filterWork As String) As Range
    ' �f�[�^�����݂���͈͂������I�Ɍ��o����
    Dim lastRow As Long

    ' ����́uF�v��Ɂu���Z�@�ցv��ʂ������Ă���̂ł��̕����ɑ΂��Č�������B
    ' �u���Z�@�ցv�̈ʒu��F���炸�ꂽ�ꍇ�A������ύX����
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    ' F��̃f�[�^���t�B���^�����O����
    ' F��̓��e��filterWork�Ɠ����s���擾����
    ws.Range("$A$1:$F$" & lastRow).AutoFilter Field:=6, Criteria1:=filterWork

    ' �t�B���^�����O���ꂽ�f�[�^���擾����
    ' Offset���\�b�h��1�s���炷���ƂŁA�w�b�_�[�s�����O����
    ' �G���[�n���h�����O��ǉ����āA�t�B���^�����O���ʂ��w�b�_�[�s�݂̂̏ꍇ���l������
    On Error Resume Next
    Set GetFilteredData = ws.Range("$A$2:$F$" & lastRow).SpecialCells(xlCellTypeVisible)
    If Err.Number <> 0 Then
        Set GetFilteredData = Nothing
    End If
    On Error GoTo 0
End Function

' ���Z�@�֖�, �w�Z�� ���󂯎���v����s�̈ꗗ��Ԃ�
Function GetFilteredDataFromFinancialInstitution(fiName As String, sheetName As String) As Range
    ' �V�[�g "sheetName" �̃^�u���J��
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' �t�B���^�����O���ꂽ�f�[�^���擾����
    Set GetFilteredDataFromFinancialInstitution = GetFilteredData(ws, fiName)

    ' �t�B���^�����O����������
    ws.AutoFilterMode = False
End Function

' ���ʂ̌������擾����
Function GetRowCount(rngData As Range) As Long
    Dim rowCnt As Long: rowCnt = 0
    Dim rng As Range
    For Each rng In rngData.Areas
        rowCnt = rowCnt + rng.EntireRow.Count
    Next
    GetRowCount = rowCnt
End Function

' �������擾����
Function GetAmount(target As String, ws As Worksheet) As Long
    GetAmount = ws.Range(target).Value
End Function


' ===========================
'  �������瓌�M��s�֌W�̏���
' ===========================

' �e���v���[�g�Ƀf�[�^��ǉ�����֐�
Sub AppendTohoDataToTemplate(filteredData As Range,lastRowTemplate As Long, amount As String, teachAmount As String, transferDate As String, wsTemplate As Worksheet, branchDict As Dictionary)
    ' �t�B���^�����O���ꂽ�f�[�^���e���v���[�g�ɒǋL����
    Dim cell As Range
    For Each cell In filteredData.Rows
        wsTemplate.Cells(lastRowTemplate, "D").Value = cell.Cells(1, "G").Value ' �������`(����)
        wsTemplate.Cells(lastRowTemplate, "E").Value = cell.Cells(1, "H").Value ' �������`(�J�i)
        wsTemplate.Cells(lastRowTemplate, "G").Value = "���M��s"
        wsTemplate.Cells(lastRowTemplate, "H").Value = cell.Cells(1, "I").Value ' �x�X��(����)
        wsTemplate.Cells(lastRowTemplate, "I").Value = branchDict(Replace(cell.Cells(1, "I").Value, "�x�X", "")) ' �x�X������x�X�ԍ����擾(�x�X����'�x�X'�̕���������΍폜)        
        wsTemplate.Cells(lastRowTemplate, "J").Value = "����"
        wsTemplate.Cells(lastRowTemplate, "K").Value = cell.Cells(1, "J").Value '
        
        ' ���H��: ���t�̏ꍇ (�w�N�� 7 ) �͋��t�����̋��z������
        If cell.Cells(1, "B").Value = "7"  Then
            wsTemplate.Cells(lastRowTemplate, "L").Value = teachAmount
        Else 
            wsTemplate.Cells(lastRowTemplate, "L").Value = amount
        End If

        wsTemplate.Cells(lastRowTemplate, "M").Value = transferDate ' �U�֓�(���M��s)
        wsTemplate.Cells(lastRowTemplate, "N").Value = cell.Cells(1, "K").Value ' �Z��

        lastRowTemplate = lastRowTemplate + 1
    Next cell
End Sub

' ���M��s�̎x�X���Ǝx�X�ԍ���R����f�[�^��ǂݍ���
Function CreateToHoBranchDictionary() As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("���M��s_�x�X���") ' �x�X���̏�����Ă���V�[�g

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 1 To lastRow
        Dim key As String
        key = ws.Cells(i, 1)

        ' �l�͎x�X�ԍ�
        Dim value As String
        value = ws.Cells(i, 3)

        ' �f�[�^�ɒǉ�
        dict.Add key, value
    Next i

    ' �R���f�[�^��Ԃ�
    Set CreateToHoBranchDictionary = dict

End Function

Sub ExecuteToho()


    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("�}�N��")

    ' �����̋��z
    Dim elementaryAmount As String ' ���w
    Dim juniorHighAmount As String ' ���w
    ' ���t���̋��z 
    Dim elementaryTeachAmount As String ' ���w
    Dim juniorHighTeachAmount As String ' ���w

    elementaryAmount = GetAmount("C9", ws)
    juniorHighAmount = GetAmount("C10", ws)

    elementaryTeachAmount = GetAmount("C11", ws)
    juniorHighTeachAmount = GetAmount("C12", ws)

    ' �U�֓�
    Dim transferDate As String 
    transferDate = GetAmount("C15", ws)


    ' �e���v���[�g�t�@�C�����J��
    Dim wbTemplate As Workbook
    Set wbTemplate = Workbooks.Open(ThisWorkbook.Path & "\templates\" &  "toho.xlsx") '�e���v���[�g�̃p�X���w�肵�Ă��������B

    ' �e���v���[�g�̍ŏ��̃V�[�g���擾����
    Dim wsTemplate As Worksheet
    Set wsTemplate = wbTemplate.Sheets(1)

    ' branchDict�����������܂�
    Dim branchDict As Dictionary
    Set branchDict = CreateToHoBranchDictionary()
    
    ' �f�[�^��ǋL����s�����Đ΂��܂��B4����Ȃ̂�1-3���w�b�_�[������ł��B
    Dim offset As Long: offset = 4

    ' ����
    Dim rngOikawa As Range
    Set rngOikawa = GetFilteredDataFromFinancialInstitution("���M", "����")
    AppendTohoDataToTemplate rngOikawa, offset, elementaryAmount, elementaryTeachAmount, transferDate, wsTemplate, branchDict

    ' ����
    Dim rngShojo As Range
    Set rngShojo = GetFilteredDataFromFinancialInstitution("���M", "����")
    AppendTohoDataToTemplate rngShojo, offset, elementaryAmount, elementaryTeachAmount, transferDate, wsTemplate, branchDict

    ' ���쒆
    Dim rngYugawa As Range
    Set rngYugawa = GetFilteredDataFromFinancialInstitution("���M", "���쒆")
    AppendTohoDataToTemplate rngYugawa, offset, juniorHighAmount, juniorHighTeachAmount, transferDate, wsTemplate, branchDict

    ' �e���v���[�g��ۑ�����
    wbTemplate.SaveAs fileName := ThisWorkbook.Path & "\result\" & "toho.xlsx"
    wbTemplate.Close savechanges := False
End Sub

' ===========================
'  �����܂œ��M��s�֌W�̏���
' ===========================


' ===========================
'  ��������JA��Ί֌W�̏���
' ===========================

' �e���v���[�g�Ƀf�[�^��ǉ�����֐�
Sub AppendJaDataToTemplate(filteredData As Range, inputDescription As String, lastRowTemplate As Long, amount As String, teachAmount As String, wsTemplate As Worksheet, branchDict As Dictionary)
    ' �t�B���^�����O���ꂽ�f�[�^���e���v���[�g�ɒǋL����
    Dim cell As Range
    For Each cell In filteredData.Rows
        wsTemplate.Cells(lastRowTemplate, "A").Value = branchDict(Replace(cell.Cells(1, "I").Value, "�x�X", "")) ' �x�X������x�X�ԍ����擾(�x�X����'�x�X'�̕���������΍폜)  
        wsTemplate.Cells(lastRowTemplate, "B").Value = cell.Cells(1, "J").Value '
        wsTemplate.Cells(lastRowTemplate, "C").Value = cell.Cells(1, "H").Value '

        ' ���t�̏ꍇ (�w�N�� 7 ) �͋��t�����̋��z������
        If cell.Cells(1, "B").Value = "7"  Then
            wsTemplate.Cells(lastRowTemplate, "D").Value = teachAmount
        Else 
            wsTemplate.Cells(lastRowTemplate, "D").Value = amount
        End If

        wsTemplate.Cells(lastRowTemplate, "E").Value = inputDescription
        lastRowTemplate = lastRowTemplate + 1
    Next cell
End Sub



' JA��΂̎x�X���Ǝx�X�ԍ���R����f�[�^��ǂݍ���
Function CreateJABranchDictionary() As Scripting.Dictionary
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("JA���_�x�X���") ' �x�X���̏�����Ă���V�[�g

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 1 To lastRow
        Dim key As String
        key = ws.Cells(i, 1)

        ' �l�͎x�X�ԍ�
        Dim value As String
        value = ws.Cells(i, 2) '�Q���

        ' �f�[�^�ɒǉ�
        dict.Add key, value
    Next i

    ' �R���f�[�^��Ԃ�
    Set CreateJABranchDictionary = dict
End Function

Sub ExecuteJA()
    Dim ws As Worksheet
    Dim inputDescription As String
     
    ' �������5�����݂Ȃǂ̒ʒ��R�����g�����
    inputDescription = InputBox("�ʒ��̃R�����g����͂��Ă�������:", "�R�����g�̓���")

    ' �������͂���̏ꍇ�A�������I������
    If inputDescription = "" Then
        MsgBox "���͂���Ă��Ȃ����߁A���u�̓L�����Z������܂����B", vbInformation, "�L�����Z��"
        Exit Sub
    End If

    ' �m�F�_�C�A���O��\��
    msgResponse = MsgBox("�ȉ��̓��e�Ő������܂����H" & vbNewLine & inputDescription, vbYesNo + vbQuestion, "�m�F")

    Set ws = ThisWorkbook.Sheets("�}�N��")
    ' �����̋��z
    Dim elementaryAmount As String ' ���w
    Dim juniorHighAmount As String ' ���w
    ' ���t���̋��z 
    Dim elementaryTeachAmount As String ' ���w
    Dim juniorHighTeachAmount As String ' ���w

    elementaryAmount = GetAmount("C9", ws)
    juniorHighAmount = GetAmount("C10", ws)

    elementaryTeachAmount = GetAmount("C11", ws)
    juniorHighTeachAmount = GetAmount("C12", ws)

    ' �����u�������v���I�����ꂽ�ꍇ
    If msgResponse = vbNo Then
        Exit Sub
    End If

    ' �e���v���[�g�t�@�C�����J��
    Dim wbTemplate As Workbook
    Set wbTemplate = Workbooks.Open(ThisWorkbook.Path & "\templates\" &  "ja.xlsx") '�e���v���[�g�̃p�X���w�肵�Ă��������B

    ' �e���v���[�g�̍ŏ��̃V�[�g���擾����
    Dim wsTemplate As Worksheet
    Set wsTemplate = wbTemplate.Sheets(1)

    ' �����ԍ��R���f�[�^
    Dim branchDict As Dictionary
    Set branchDict = CreateJABranchDictionary()
    
    ' �f�[�^��ǋL����s�����Đ΂��܂��B2����Ȃ̂�1-3���w�b�_�[������ł��B
    Dim offset As Long: offset = 2

    ' ����
    Dim rngOikawa As Range
    Set rngOikawa = GetFilteredDataFromFinancialInstitution("JA", "����")
    If Not rngOikawa Is Nothing Then
        AppendJaDataToTemplate rngOikawa, inputDescription, offset, elementaryAmount, elementaryTeachAmount, wsTemplate, branchDict
    End If
    Set rngOikawa = GetFilteredDataFromFinancialInstitution("�i�`��Â��", "����")
    If Not rngOikawa Is Nothing Then
        AppendJaDataToTemplate rngOikawa, inputDescription, offset, elementaryAmount, elementaryTeachAmount, wsTemplate, branchDict
    End If

    ' ����
    Dim rngShojo As Range
    Set rngShojo = GetFilteredDataFromFinancialInstitution("JA", "����")
    If Not rngShojo Is Nothing Then
        AppendJaDataToTemplate rngShojo, inputDescription, offset, elementaryAmount, elementaryTeachAmount, wsTemplate, branchDict
    End If
    Set rngShojo = GetFilteredDataFromFinancialInstitution("�i�`��Â��", "����")
    If Not rngShojo Is Nothing Then
        AppendJaDataToTemplate rngShojo, inputDescription, offset, elementaryAmount, elementaryTeachAmount, wsTemplate, branchDict
    End If

    ' ���쒆
    Dim rngYugawa As Range
    Set rngYugawa = GetFilteredDataFromFinancialInstitution("JA", "���쒆")
    If Not rngYugawa Is Nothing Then
        AppendJaDataToTemplate rngYugawa, inputDescription, offset, juniorHighAmount, juniorHighTeachAmount, wsTemplate, branchDict
    End If
    Set rngYugawa = GetFilteredDataFromFinancialInstitution("�i�`��Â��", "���쒆")
    If Not rngYugawa Is Nothing Then
        AppendJaDataToTemplate rngYugawa, inputDescription, offset, juniorHighAmount, juniorHighTeachAmount, wsTemplate, branchDict
    End If

    ' �e���v���[�g��ۑ�����
    wbTemplate.SaveAs fileName := ThisWorkbook.Path & "\result\" & "ja.xlsx"
    wbTemplate.Close savechanges := False
End Sub




' ===========================
'  ��������w�N�ڍs�̏���
' ===========================

Sub ExecuteMigration()
    Dim inputNumber As String
    ' �V�����̐l�������
    inputNumber = InputBox("�V�����̐l������͂��Ă�������:", "�V�����̐l��")

    ' �������͂���̏ꍇ�A�������I������
    If inputNumber = "" Then
        MsgBox "���͂���Ă��Ȃ����߁A���u�̓L�����Z������܂����B", vbInformation, "�L�����Z��"
        Exit Sub
    End If

    ' �m�F�_�C�A���O��\��
    msgResponse = MsgBox("�ȉ��̓��e�Ŋw�N���X�V���܂����H" & vbNewLine & inputNumber, vbYesNo + vbQuestion, "�m�F")
    
    ' �����u�������v���I�����ꂽ�ꍇ
    If msgResponse = vbNo Then
        Exit Sub
    End If

    Dim i As Long

    '�V�[�g��ݒ�
    Dim wsOikawa As Worksheet, wsShojo As Worksheet, wsYugawa As Worksheet
    Set wsOikawa = ThisWorkbook.Sheets("����")
    Set wsShojo = ThisWorkbook.Sheets("����")
    Set wsYugawa = ThisWorkbook.Sheets("���쒆")

    '�Ō�̍s���擾
    Dim LastRowOikawa As Long
    Dim LastRowShojo As Long
    Dim LastRowYugawa As Long
    LastRowOikawa = wsOikawa.Cells(wsOikawa.Rows.Count, "A").End(xlUp).Row
    LastRowShojo = wsShojo.Cells(wsShojo.Rows.Count, "A").End(xlUp).Row
    LastRowYugawa = wsYugawa.Cells(wsYugawa.Rows.Count, "A").End(xlUp).Row

    ' ===========================
    '  �������瓒�쒆�̏���
    ' ===========================

    '���쒆�A�w�N��3�̍s���폜
    For i = LastRowYugawa To 2 Step -1 ' 2�s�ڂ���J�n
        If wsYugawa.Cells(i, 2).Value = 3 Then
            wsYugawa.Rows(i).Delete
        End If
    Next i

    '���쒆�A�w�N�̍X�V
    For i = 2 To LastRowYugawa
        If wsYugawa.Cells(i, 2).Value = 1 Then
            wsYugawa.Cells(i, 2).Value = 2
        ElseIf wsYugawa.Cells(i, 2).Value = 2 Then
            wsYugawa.Cells(i, 2).Value = 3
        End If
    Next i

    ' ===========================
    '  �������狈�쏬�̏���
    ' ===========================

    '���쏬����w�N��6�̍s�𓒐쒆�V�[�g�̐擪�ɃR�s�[
    For i = LastRowOikawa To 2 Step -1 ' 2�s�ڂ���J�n
        If wsOikawa.Cells(i, 2).Value = 6 Then ' �w�N���U�̂���
            '�R�s�[����͈͂�ݒ�
            Set rngToCopy = wsOikawa.Rows(i)
            '�}������ʒu��ݒ�
            Set rngDest = wsYugawa.Rows(2)
            '�s��}��
            rngDest.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            '�f�[�^���R�s�[
            rngToCopy.Copy Destination:=wsYugawa.Rows(2)
            '����V�[�g����s���폜
            wsOikawa.Rows(i).Delete
        End If
    Next i

    '���쏬�A�w�N�̍X�V
    For i = 2 To LastRowOikawa
        If wsOikawa.Cells(i, 2).Value = 1 Then
            wsOikawa.Cells(i, 2).Value = 2
        ElseIf wsOikawa.Cells(i, 2).Value = 2 Then
            wsOikawa.Cells(i, 2).Value = 3
        ElseIf wsOikawa.Cells(i, 2).Value = 3 Then
            wsOikawa.Cells(i, 2).Value = 4
        ElseIf wsOikawa.Cells(i, 2).Value = 4 Then
            wsOikawa.Cells(i, 2).Value = 5
        ElseIf wsOikawa.Cells(i, 2).Value = 5 Then
            wsOikawa.Cells(i, 2).Value = 6
        End If
    Next i

    '�擪��20�s�̋󔒂�ǉ�
    wsOikawa.Rows("2:" & inputNumber).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

    ' ===========================
    '  �������珟�포�̏���
    ' ===========================

    '���포����w�N��6�̍s�𓒐쒆�V�[�g�̐擪�ɃR�s�[
    For i = LastRowShojo To 2 Step -1 ' 2�s�ڂ���J�n
        If wsShojo.Cells(i, 2).Value = 6 Then ' �w�N���U�̂���
            '�R�s�[����͈͂�ݒ�
            Set rngToCopy = wsShojo.Rows(i)
            '�}������ʒu��ݒ�
            Set rngDest = wsYugawa.Rows(2)
            '�s��}��
            rngDest.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            '�f�[�^���R�s�[
            rngToCopy.Copy Destination:=wsYugawa.Rows(2)
            '����V�[�g����s���폜
            wsShojo.Rows(i).Delete
        End If
    Next i

    '���포�A�w�N�̍X�V
    For i = 2 To LastRowShojo
        If wsShojo.Cells(i, 2).Value = 1 Then
            wsShojo.Cells(i, 2).Value = 2
        ElseIf wsShojo.Cells(i, 2).Value = 2 Then
            wsShojo.Cells(i, 2).Value = 3
        ElseIf wsShojo.Cells(i, 2).Value = 3 Then
            wsShojo.Cells(i, 2).Value = 4
        ElseIf wsShojo.Cells(i, 2).Value = 4 Then
            wsShojo.Cells(i, 2).Value = 5
        ElseIf wsShojo.Cells(i, 2).Value = 5 Then
            wsShojo.Cells(i, 2).Value = 6
        End If
    Next i

    '���쒆�A6�N����1�N����
    For i = 2 To LastRowYugawa
        If wsYugawa.Cells(i, 2).Value = 6 Then
            wsYugawa.Cells(i, 2).Value = 1
        End If
    Next i


    '�擪��20�s�̋󔒂�ǉ�
    wsShojo.Rows("2:" & inputNumber).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromRightOrBelow

   '�ʖ��ŕۑ�
    ThisWorkbook.SaveAs fileName := ThisWorkbook.Path & "\new_" & ThisWorkbook.Name

End Sub