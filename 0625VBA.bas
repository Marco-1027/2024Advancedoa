Attribute VB_Name = "Module1"
Sub FirstVBAHelp()
'
' FirstVBAHelp ����
'
' �ֳt��: Ctrl+q

   '�I��A1�x�s��
    Range("A1").Select
   '�I��A1�_�l�̥��CCTRL+���U
    Selection.End(xlDown).Select
    '�u��������r�n�����޸��]�_�� Row���N��O�C���X������
    MsgBox "���i�D��,�C�Ʀ�" & Selection.End(xlDown).Rows.Row
    '�I��̥���CTRL+���k
    Selection.End(xlToRight).Select
    '�u��������̥k��,����� ��r�n�����޸��]�_�� Couumns�O�涰�X������
    MsgBox "���i�D��,��Ʀ�" & Selection.End(xlToRight).Columns.Column
    '���槹��,��Ц^��A1
    Range("A1").Select
    
End Sub
