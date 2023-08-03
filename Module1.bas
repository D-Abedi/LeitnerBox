Attribute VB_Name = "Module1"
Public i As Integer, j As Integer, Flag As Boolean
Sub AddVocabForm()
    AddVocab.Show
End Sub
Sub ReviewForm()
On Error Resume Next
    i = 1
    Review.Show
End Sub
Sub looper(i As Integer)
UserNamei = Application.UserName
With Workbooks("Vocab.xlsm").Worksheets("Sheet1").ListObjects("tblVocab")
nTblVocab = .ListRows.Count
    For i = i To nTblVocab
        If .ListColumns("Review Date").DataBodyRange(i).Value <= Now Then
            Review.boxWord.Value = .ListColumns("Word").DataBodyRange(i).Value
            Review.boxPoS.Value = .ListColumns("Pos").DataBodyRange(i).Value
            Exit Sub
        End If
    Next i
    EarlyDate = WorksheetFunction.Min(.ListColumns("Review Date").DataBodyRange)
MsgBox "Dear " & UserNamei & "!" & vbCrLf & vbCrLf & _
        "You did a great job. There is no word to review on this turn." & vbCrLf & _
        "Your next turn will be on:  " & Format(EarlyDate, "ddd mmm/dd, hh:mm."), , "Review Finished"
Unload Review
End With
End Sub
Sub blanker(FormName As UserForm)
    FormName.boxWord.Value = ""
    FormName.boxPoS.Value = ""
    FormName.boxSyn.Value = ""
    FormName.boxPeTr.Value = ""
    FormName.boxDef.Value = ""
    FormName.boxExample.Value = ""
End Sub
'------
Private Sub MyRightClickMenu()
    Application.CommandBars("Cell").Reset
    Dim cbc As CommandBarControl
    For Each cbc In Application.CommandBars("Cell").Controls
        cbc.Visible = False
    Next cbc
    CapArr = Array("Cut", "Copy", "Paste", "Select All", "Delete")
    ActArr = Array("iCut", "iCopy", "iPaste", "iAll", "iDel")
    For i = 0 To UBound(CapArr)
        With Application.CommandBars("Cell").Controls.Add(temporary:=True)
            .Caption = CapArr(i)
            .OnAction = ActArr(i)
        End With
    Next i
    Application.CommandBars("Cell").ShowPopup
End Sub
Private Sub iCut()
SendKeys "^x"
End Sub
Private Sub iCopy()
SendKeys "^c"
End Sub
Private Sub iPaste()
SendKeys "^v"
End Sub
Private Sub iAll()
SendKeys "^a"
End Sub
Private Sub iDel()
SendKeys "{DEL}"
End Sub
'------------------------------





