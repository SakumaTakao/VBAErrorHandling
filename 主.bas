Attribute VB_Name = "��"
'Copyright 2022-2024 SAKUMA, Takao
Option Explicit

Public Type ��
  ID As Long
  Source As String
  Description As String
  HelpFile As String
  HelpContext As Long
  LastDllError As Long
  
  Erl As Long '�s���x���̐ݒ�\�l��0�`2147483647�iLong�^�̂������łȂ��l�j
  When As Date
End Type

Public Enum ErrNumber
  �����Y���� = 5& 'Collection��Dictionary�̕�����L�[�ɊY����
  �W���� = 5&
  �p�X���󕶎� = 5&
  �I�[�o�[�t���[ = 6&
  �ԊO = 9& '�z��ACollection�ADictionary�̃C���f�N�X���͈͊O
  ReDim���� = 9&
  �^�Ⴂ = 13&
  �t�@�C�����s�� = 52&
  �t�@�C�����쐬 = 53&
  �t�@�C�����J = 55&
  �t�@�C���ǎ��p = 70&
  �t�@�C���p�X���� = 75&
  �t�@�C���p�X�s�� = 76&
  ������ = 91&
  �C���X�^���X���ݒ� = 91&
  Null�s�� = 94&
  �C���X�^���X�v = 424&
  ����String�ɃC���X�^���X���� = 438&
  ���\�b�h���� = 438&
  �������@ = 438&
  �������s = 440&
  �����o�^�� = 457&
  
  ��� = 65535
End Enum

Public Enum ErrID��65536�ȏ�
  �ďo��G���[ = ErrNumber.��� + 1&
  �p�@����
  ��������
End Enum

Private Const CodeName As String = "��"
Private Const LastDllError�A���� As String = " LastDllError = "

Private ErrCopy As ��
Private ���̏����l As ��
Private �a�ނŔ����������G���[ As Boolean

Private �^�C���X�^���X As New �^

Public Sub �^��(ByVal �����v���V�[�W�����Ȃ� As String, Optional ByVal �q As String)
  Dim �G���[ As �G���[�̕ۑ��ƕ���
  Set �G���[ = New �G���[�̕ۑ��ƕ���
  
  On Error GoTo OnError
  Dim �� As ��
  With �G���[.�G���[
    ��.Source = IIf(�����v���V�[�W�����Ȃ� = "", .Source, �����v���V�[�W�����Ȃ�)
    ��.Description = IIf(�q = "", .Description, �q)

    If .ID Then
      ��.Description = CStr(.ID) & ": " & ��.Description
      If .Erl Then ��.Source = ��.Source & " ���x��" & CStr(Erl())
    End If
  End With
  ��.When = Now()

  �^�C���X�^���X.�^�� ��
Exit Sub
OnError:
  �^�� CodeName & ".�^��"
  Resume Next
End Sub

Public Sub �a��(ByVal ErrID As ErrID��65536�ȏ�, Optional ByVal �����v���V�[�W�����Ȃ� As String, Optional ByVal �q As String)
  If ErrID <= ErrNumber.��� Then
    �^�� �����v���V�[�W�����Ȃ�, "�u�a�ށv�̈����uErrID�v�ɂ�65536�ȏ�̐�������������܂��� " & ErrID & " ���^�����܂����B"
    ErrID = -ErrID
    If ErrID = 0 Then ErrID = ErrID��65536�ȏ�.��������
  End If
  
  Dim �� As ��
  With ��
    .ID = -ErrID
    .Source = �����v���V�[�W�����Ȃ�
    .Description = �q
    
    'Err().Raise�̍ۂɈ���HelpFile���ȗ�����ƃf�t�H���g�l���ݒ肳��Ă��܂�
    '�����h�����ߋ󕶎���ݒ肷��
    .HelpFile = ""
    
    '����HelpContext��0��^�����Err().Raise�̍ۂ�1000440�ɕϊ�����Ă��܂��̂ŁA-1��^����
    .HelpContext = -1
  End With
  �G���[() = ��
  �a�ނŔ����������G���[ = True
  
  '�{�v���V�[�W����Optional�������ȗ�����Ă����ꍇ�́AErr().Raise�̍ۂɑΉ�����������ȗ����f�t�H���g�l��ݒ肳����
  With Err() '���O�́u�G���[() = ���v�ɂ��X�V����Ă���
    If �����v���V�[�W�����Ȃ� = "" Then
      If �q = "" Then
        .Raise .Number, , , .HelpFile, .HelpContext
      Else
        .Raise .Number, , .Description, .HelpFile, .HelpContext
      End If
    Else
      If �q = "" Then
        .Raise .Number, .Source, , .HelpFile, .HelpContext
      Else
        .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
      End If
    End If
  End With
End Sub

Public Sub �^���a�ށ�OnError(ByVal �����v���V�[�W�����Ȃ� As String)
  With �G���[()
    If .ID Then
      �^�� �����v���V�[�W�����Ȃ�
      If .ID < -ErrNumber.��� And .HelpFile = "" Then '���ꂪTrue�ɂȂ邱�Ƃ͖����H
        �a�� -.ID
      Else
        �a�� �ďo��G���[, �����v���V�[�W�����Ȃ�, "�u" & �����v���V�[�W�����Ȃ� & "�v�ŃG���[�𐶂��܂����B"
      End If
    Else
      �^�� �����v���V�[�W�����Ȃ�, "�G���[�������Ă��Ȃ��̂Ɂu�^���a�ށ�OnError�v���ĂԂ��Ƃ͋�����܂���B"
    End If
  End With
End Sub

Public Property Set �^(ByVal �^ As �^)
  On Error GoTo OnError
  If �^ Is Nothing Then Set �^ = New �^
  Set �^�C���X�^���X = �^
Exit Property
OnError:
  �^���a�ށ�OnError "Set �^"
End Property

Public Property Get �G���[() As ��
'  If (ErrCopy.ID <> 0) And (Err().Number <> 0) Then
  If CBool(ErrCopy.ID) * Err().Number Then
    With ErrCopy
      �G���[.ID = .ID
      �G���[.Source = .Source
      �G���[.Description = .Description
      �G���[.HelpFile = .HelpFile
      �G���[.HelpContext = .HelpContext
      
      If �a�ނŔ����������G���[ Then
        �G���[.LastDllError = Err().LastDllError
        �G���[.Erl = Erl()
        �a�ނŔ����������G���[ = False
      Else
        �G���[.LastDllError = .LastDllError
        �G���[.Erl = .Erl
      End If
    End With
  Else
    With Err()
      �G���[.ID = .Number
      �G���[.Source = .Source
      �G���[.Description = .Description
      �G���[.HelpFile = .HelpFile
      �G���[.HelpContext = .HelpContext
      �G���[.LastDllError = .LastDllError
      �G���[.Erl = Erl()
    End With
  End If
  ErrCopy = ���̏����l
End Property

Public Property Let �G���[(�� As ��)
  With ErrCopy
    .ID = ��.ID
    .Source = ��.Source
    .Description = ��.Description
    .HelpFile = ��.HelpFile
    '����HelpContext��0��^�����Err().Raise����1000440�ɕϊ�����Ă��܂��̂ŁA-1��^����
    .HelpContext = IIf(��.HelpContext, -1, ��.HelpContext)
    .LastDllError = ��.LastDllError
    .Erl = ��.Erl

    Err().Number = IIf(.ID > ErrNumber.���, -.ID, .ID)
    Err().Source = .Source
    Err().Description = .Description
    Err().HelpFile = .HelpFile
    Err().HelpContext = .HelpContext
    
    'Err().LastDllError�͓ǎ��p�Ȃ̂ŏ���������ׂ��Ƃ���Err().Description�ɒǋL
    If Err().LastDllError <> .LastDllError _
      Then Err().Description = .Description & LastDllError�A���� & .LastDllError
    
  End With
End Property


