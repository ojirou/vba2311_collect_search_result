Attribute VB_Name = "Module1"
Option Explicit
Sub CollectSearchResult()
Dim fso As FileSystemObject
Set fso = New FileSystemObject
Dim pass As String
'pass = "\C:Users\user\git\excel_vba\sample_google_search"
'pass = ThisWorkbook.Path & "\test"
pass = ThisWorkbook.Path & "\sample_google_search"
Dim FileName As String
Dim i As Long, j As Long '�������ʂ̍s��,�������ʏW�v�̍s��
j = 2
Dim f As File
For Each f In fso.GetFolder(pass).Files
        ' �t�@�C��������g���q���폜���� FileName �ɑ��
    With Workbooks.Open(f)
        With .Worksheets(1)
        FileName = fso.GetBaseName(f.Name)
            i = 2
            Do While .Cells(i, 1).Value <> ""
                Sheet1.Cells(j, 1).Value = FileName            '�t�@�C����
                Sheet1.Cells(j, 2).Value = .Cells(i, 1).Value  'ID
                Sheet1.Cells(j, 3).Value = .Cells(i, 2).Value   '�^�C�g��
                Sheet1.Cells(j, 4).Value = .Cells(i, 3).Value   'URL
                i = i + 1
                j = j + 1
            Loop
        End With
        .Close
    End With
Next f
End Sub
