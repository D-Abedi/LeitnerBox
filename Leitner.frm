VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Leitner 
   Caption         =   "Leitner"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "Leitner.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Leitner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub boxAnt_Enter()
    With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
        Me.boxAnt.Value = .ListColumns("Ant.").DataBodyRange(i).Value
    End With
End Sub
Private Sub boxSyn_Enter()
    With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
        Me.boxSyn.Value = .ListColumns("Syn.").DataBodyRange(i).Value
    End With
End Sub
Private Sub boxDefinition_Enter()
    With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
        Me.boxDefinition.Value = .ListColumns("Definition").DataBodyRange(i).Value
    End With
End Sub
Private Sub boxExample_Enter()
    With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
        Me.boxExample.Value = .ListColumns("Example").DataBodyRange(i).Value
    End With
End Sub
Private Sub UserForm_Initialize()
    Call looper(i)
End Sub
Private Sub btnAnswer_Click()
    With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
        Me.boxDefinition.Value = .ListColumns("Definition").DataBodyRange(i).Value
        Me.boxSyn.Value = .ListColumns("Syn.").DataBodyRange(i).Value
        Me.boxAnt.Value = .ListColumns("Ant.").DataBodyRange(i).Value
        Me.boxExample.Value = .ListColumns("Example").DataBodyRange(i).Value
    End With
End Sub
Private Sub btnClose_Click()
    RowCounter = 0
    i = 0
    Unload Leitner
End Sub
Private Sub btnFalse_Click()
Call blanker(Leitner)
With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
    .ListColumns("Review Date").DataBodyRange(i).Value = Date
    .ListColumns("Step").DataBodyRange(i).Value = 0
End With
i = i + 1
Call looper(i)
End Sub
Private Sub btnTrue_Click()
Call blanker(Leitner)
With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
    .ListColumns("Review Date").DataBodyRange(i).Value = Date + (2 ^ .ListColumns("Step").DataBodyRange(i).Value)
    .ListColumns("Step").DataBodyRange(i).Value = .ListColumns("Step").DataBodyRange(i).Value + 1
End With
Call looper(i)
End Sub
'------------
