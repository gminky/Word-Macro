VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �g�\���ؑ�form 
   Caption         =   "�c�[��"
   ClientHeight    =   2064
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   5496
   OleObjectBlob   =   "�g�\���ؑ�form.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�g�\���ؑ�form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub �s����_�X�^�C���{�^��_Click()
    If Selection.Paragraphs.FarEastLineBreakControl = True Then
        Selection.Paragraphs.FarEastLineBreakControl = False
        �g�\���ؑ�form.�s����_�X�^�C���{�^��.Caption = "�ݒ肷��"
    Else
        Selection.Paragraphs.FarEastLineBreakControl = True
        �g�\���ؑ�form.�s����_�X�^�C���{�^��.Caption = "�ݒ肵�Ȃ�"
    End If
End Sub

Private Sub �\���{�^��_Click()
    If �\���{�^��.Caption = "�\������" Then
        '��\���ɂ���ꍇ
        �\���{�^��.Caption = "�\�����Ȃ�"
        Call �t�c���c�p�g�\��
    Else
        '�\������ꍇ
        �\���{�^��.Caption = "�\������"
        Call �t�c���c�p�g��\��
    
    End If
End Sub
