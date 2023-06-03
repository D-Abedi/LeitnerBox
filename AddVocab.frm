VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddVocab 
   Caption         =   "Add Vocabulary to Leitner Box"
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
Private Sub boxAnt_Enter()
    If Me.boxAnt.Text = "Antonyms" Then
        Me.boxAnt.ForeColor = RGB(0, 0, 0)
        Me.boxAnt.Text = ""
    End If
End Sub
Private Sub boxAnt_AfterUpdate()
    If Me.boxAnt.Text = "" Then
        Me.boxAnt.ForeColor = RGB(109, 109, 109)
        Me.boxAnt.Text = "Antonyms"
    End If
End Sub
Private Sub boxDefinition_Enter()
    If Me.boxDefinition.Text = "Definition" Then
        Me.boxDefinition.ForeColor = RGB(0, 0, 0)
        Me.boxDefinition.Text = ""
    End If
End Sub
Private Sub boxDefinition_AfterUpdate()
    If Me.boxDefinition.Text = "" Then
        Me.boxDefinition.ForeColor = RGB(109, 109, 109)
        Me.boxDefinition.Text = "Definition"
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
    txtArray = Array("New Word", "Part of Speech", "Synonyms", "Antonyms", "Definition", "Examples")
    tblArray = Array("Word", "PoS", "Syn.", "Ant.", "Definition", "Example")
    With AddVocab
        boxArray = Array(.boxWord, .boxPoS, .boxSyn, .boxAnt, .boxDefinition, .boxExample)
    End With
    For i = 0 To UBound(boxArray)
        If boxArray(i).Text = txtArray(i) Then EmptyList = EmptyList + "  •  " & txtArray(i) & vbCrLf
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
        .ListColumns("Review Date").DataBodyRange(.ListRows.Count).Value = Date
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


