Attribute VB_Name = "�L"
'Copyright 2022-2024 SAKUMA, Takao
Option Explicit

Public Enum �q
  �����@�j�nEnum�A�m��`�l�V�J���T���}�Z���K�B�K�^�G�����}�V�^  ' = &H80000000
  �����@�m�l���A�j���L���G�e���s�V�}�X
  �����@�jNothing���^�G���R�g�n���T���}�Z��
  �@�jNothing���^�G���R�g�n���T���}�Z�� 'Set Property�̏ꍇ
  �@�f����[�����W�}�V�^
  �@�j�^�G�^�����K�s���f�X
  Class�@��New�f�����X���R�g�n���T���}�Z��
  �@���g�b�e�N�_�T�C
End Enum

Private Const CodeName As String = "�L"

Private Const �@ As Long = 9312 'ChrW("�@")
Private Const �� As String = vbCr
Private Const �q�̘A�� As String = _
    "�����u�@�v�ɂ�Enum�u�A�v�̒�`�l����������܂��� �B ���^�����܂����B" & �� & _
    "�����u�@�v�̒l�� �A �ɏ��������đ��s���܂��B" & �� & _
    "�����u�@�v��Nothing��^���邱�Ƃ͋�����܂���B" & �� & _
    "�u�@�v��Nothing��^���邱�Ƃ͋�����܂���B" & �� & _
    "�u�@�v�ŃG���[�𐶂��܂����B" & �� & _
    "�u�@�v�ɗ^�����������s���ł��B" & �� & _
    "Class�u�@�v��New�Ő������邱�Ƃ͋�����܂���B" & �� & _
    "�u�@�v���g���Ă��������B"

Private �q�� As Variant '�q��() As String

Public Function �q��(ByVal Enum�q As �q, ParamArray �@���̒l() As Variant) As String
  Dim �G���[ As �G���[�̕ۑ��ƕ���
  Set �G���[ = New �G���[�̕ۑ��ƕ���
  
  On Error GoTo OnError
1:
  �q�� = �q��(Enum�q)
2:
  Dim i As Long
  For i = LBound(�@���̒l) To UBound(�@���̒l)
    �q�� = Replace(�q��, ChrW(�@ + i), �@���̒l(i))
  Next
Exit Function
OnError:
  With Err()
    Select Case .Number
      Case ErrNumber.�^�Ⴂ, ErrNumber.Null�s��, ErrNumber.�C���X�^���X���ݒ�
        If Erl() = 1 Then
          �q�� = Split(�q�̘A��, ��)
          Resume
        End If
        �^�� "�L��.�L.�q��", "�����u�@���̒l�v��" & CStr(i + 1&) & "�Ԗڂ̗v�f��String�^�ɕϊ��ł��܂���B"
        Resume Next
      Case ErrNumber.�ԊO
        �^�� "�L��.�L.�q��", "�����uEnum�q�v�ɊY�����镶���񂪓o�^����Ă��܂���B"
        �q�� = "�iFunction�u�q�ԁv���G���[�ɂȂ�܂����B"
      Case Else
         �^�� "�L��.�L.�q��"
        If �q�� = "" Then �q�� = "�iFunction�u�q�ԁv���G���[�ɂȂ�܂����B�j"
    End Select
  End With
End Function

