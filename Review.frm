VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Review 
   Caption         =   "Review Words"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   OleObjectBlob   =   "Review.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Review"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub boxPeTr_Enter()
    With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
        Me.boxPeTr.Value = .ListColumns("PeTr").DataBodyRange(i).Value
    End With
End Sub
Private Sub boxSyn_Enter()
    With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
        Me.boxSyn.Value = .ListColumns("Syn.").DataBodyRange(i).Value
    End With
End Sub
Private Sub boxDef_Enter()
    With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
        Me.boxDef.Value = .ListColumns("Definition").DataBodyRange(i).Value
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
        Me.boxDef.Value = .ListColumns("Definition").DataBodyRange(i).Value
        Me.boxSyn.Value = .ListColumns("Syn.").DataBodyRange(i).Value
        Me.boxPeTr.Value = .ListColumns("PeTr").DataBodyRange(i).Value
        Me.boxExample.Value = .ListColumns("Example").DataBodyRange(i).Value
    End With
End Sub
Private Sub btnClose_Click()
    RowCounter = 0
    i = 0
    WMP1.Close
    Unload Review
End Sub
Private Sub btnFalse_Click()
Call blanker(Review)
With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
    .ListColumns("Review Date").DataBodyRange(i).Value = Now + TimeValue("00:30:00")
    .ListColumns("Step").DataBodyRange(i).Value = 0
End With
i = i + 1
Call looper(i)
End Sub
Private Sub btnTrue_Click()
Call blanker(Review)
With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
    .ListColumns("Review Date").DataBodyRange(i).Value = Date + (2 ^ .ListColumns("Step").DataBodyRange(i).Value)
    .ListColumns("Step").DataBodyRange(i).Value = .ListColumns("Step").DataBodyRange(i).Value + 1
End With
Call looper(i)
End Sub
Private Sub Listen_Click()
Dim iURL As String
'--- Check if Longman Dictionary's pronunciation is available. Otherwise, Google pronunciation plays
With Workbooks("Vocab.xlsm").Worksheets("sheet1").ListObjects("tblVocab")
    iURL = "https://www.ldoceonline.com/media/english/ameProns/" & .ListColumns("Word").DataBodyRange(i).Value & ".mp3"
    If URLExists(iURL) = False Then
        iURL = "https://ssl.gstatic.com/dictionary/static/sounds/20220808/" & _
        LCase(.ListColumns("Word").DataBodyRange(i).Value) & "--_gb_1.mp3"
    End If
    WMP1.url = iURL
End With
End Sub
'------------
Function URLExists(url As String) As Boolean
    Dim Request As Object
    Dim rc As Variant
    On Error GoTo EndNow
    Set Request = CreateObject("WinHttp.WinHttpRequest.5.1")
    With Request
      .Open "GET", url, False
      .Send
      rc = .StatusText
    End With
    Set Request = Nothing
    If rc = "OK" Then URLExists = True
    Exit Function
EndNow:
End Function