VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddVocab 
   Caption         =   "Add a New Word"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "AddVocab.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddVocab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text 'Enabled to NOT case-sensitive search for strings
Private Sub boxWord_Change()
    With Workbooks("Vocab.xlsm").Worksheets("Sheet1").ListObjects("tblVocab")
        WordList = .ListColumns("Word").DataBodyRange.Value
        For Each Item In WordList
            If Me.boxWord.Value = Item Then
                MsgBox "You have this word on your LeitnerBox.", vbInformation, "Duplicate Word"
                With Me.boxWord
                    .SetFocus
                    .SelStart = 0
                    .SelLength = Len(.Text)
                End With
                Exit Sub
            End If
        Next Item
    End With
End Sub
Private Sub boxWord_Enter()
    If Me.boxWord.Text = "New Word" Then
        Me.boxWord.ForeColor = RGB(0, 0, 0)
        Me.boxWord.Text = ""
    End If
End Sub
Private Sub boxWord_AfterUpdate()
    If Me.boxWord.Text = "" Then
        Me.boxWord.ForeColor = RGB(109, 109, 109)
        Me.boxWord.Value = "New Word"
    End If
End Sub
Private Sub boxPos_Enter()
    If Me.boxPoS.Text = "Part of Speech" Then
        Me.boxPoS.ForeColor = RGB(0, 0, 0)
        Me.boxPoS.Text = ""
    End If
End Sub
Private Sub boxPos_AfterUpdate()
    If Me.boxPoS.Text = "" Then
        Me.boxPoS.ForeColor = RGB(109, 109, 109)
        Me.boxPoS.Text = "Part of Speech"
    End If
End Sub
Private Sub boxSyn_Enter()
    If Me.boxSyn.Text = "Synonyms" Then
        Me.boxSyn.ForeColor = RGB(0, 0, 0)
        Me.boxSyn.Text = ""
    End If
End Sub
Private Sub boxSyn_AfterUpdate()
    If Me.boxSyn.Text = "" Then
        Me.boxSyn.ForeColor = RGB(109, 109, 109)
        Me.boxSyn.Text = "Synonyms"
    End If
End Sub
Private Sub boxPeTr_Enter()
SendKeys "%+"
    If Me.boxPeTr.Text = " —Ã„Â" Then
        Me.boxPeTr.ForeColor = RGB(0, 0, 0)
        Me.boxPeTr.Text = ""
    End If
End Sub
Private Sub boxPeTr_AfterUpdate()
    If Me.boxPeTr.Text = "" Then
        Me.boxPeTr.ForeColor = RGB(109, 109, 109)
        Me.boxPeTr.Text = " —Ã„Â"
    End If
SendKeys "%+"
End Sub
Private Sub boxDef_Enter()
    If Me.boxDef.Text = "Definition" Then
        Me.boxDef.ForeColor = RGB(0, 0, 0)
        Me.boxDef.Text = ""
    End If
End Sub
Private Sub boxDef_AfterUpdate()
    If Me.boxDef.Text = "" Then
        Me.boxDef.ForeColor = RGB(109, 109, 109)
        Me.boxDef.Text = "Definition"
    End If
End Sub
Private Sub boxExample_Enter()
    If Me.boxExample.Text = "Examples" Then
        Me.boxExample.ForeColor = RGB(0, 0, 0)
        Me.boxExample.Text = ""
    End If
End Sub
Private Sub boxExample_AfterUpdate()
    If Me.boxExample.Text = "" Then
        Me.boxExample.ForeColor = RGB(109, 109, 109)
        Me.boxExample.Text = "Examples"
    End If
End Sub
Private Sub btnAddWord_Click()
    Dim EmptyList As String
    txtArray = Array("New Word", "Part of Speech", "Synonyms", " —Ã„Â", "Definition", "Examples")
    tblArray = Array("Word", "PoS", "Syn.", "PeTr", "Definition", "Example")
    With AddVocab
        boxArray = Array(.boxWord, .boxPoS, .boxSyn, .boxPeTr, .boxDef, .boxExample)
    End With
    For i = 0 To UBound(boxArray)
        If boxArray(i).Text = txtArray(i) Then EmptyList = EmptyList + "  ï  " & txtArray(i) & vbCrLf
    Next i
    If EmptyList <> "" Then
        answer = MsgBox("The list below is the fields that are left empty:" _
        & vbCrLf & EmptyList & vbCrLf & _
        "Do you want to fill them up?", vbQuestion + vbYesNo, "Empty Field")
        If answer = vbYes Then Exit Sub
    End If
    With Workbooks("Vocab.xlsm").Worksheets("Sheet1").ListObjects("tblVocab")
        .ListRows.Add
        .ListColumns("Step").DataBodyRange(.ListRows.Count).Value = 0
        .ListColumns("Review Date").DataBodyRange(.ListRows.Count).Value = Now + TimeValue("00:30:00")
        For i = 0 To UBound(boxArray)
            If boxArray(i).Text = txtArray(i) Then boxArray(i).Text = ""
            .ListColumns(tblArray(i)).DataBodyRange(.ListRows.Count).Value = _
            boxArray(i).Text
        Next i
    End With
    Unload AddVocab
End Sub
Private Sub btnClose_Click()
    Unload AddVocab
End Sub
'--------
Private Sub boxExample_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
If Button = 2 Then Run "MyRightClickMenu"
End Sub
Private Sub UserForm_Terminate()
Application.CommandBars("Cell").Reset
End Sub


