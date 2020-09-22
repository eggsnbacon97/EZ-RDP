Attribute VB_Name = "Module1"
Option Compare Text
Sub Remote_Client_Server()
RDPWindow = Shell("C:\windows\system32\mstsc.exe /v:" & Range("B3"), 1)
End Sub
Sub Clear_Field()
Dim clear As String
clear = ""
Range("B3").Value = clear
End Sub
Sub Search()
Dim room As Integer

bldg = InputBox("Which building?", ["Search"])
room = InputBox("Which room?", ["Search"])

    For i = 7 To 90
        For j = 1 To 5
        
        If Cells(i, j).Value = (room) And Cells(i, 5).Value = (bldg) Then GoTo Found
        
        Next j
      Next i


    
NotFound:
    MsgErr = "Not Found"
    MsgBox Prompt:=MsgErr, Title:="Search"
Exit Sub

Found:
    MsgBox (Cells(i, 4).Value)
    Range("B3").Value = (Cells(i, 4).Value)
End Sub
