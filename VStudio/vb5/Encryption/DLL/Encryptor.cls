VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Encryptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Result As String
Private Code As Integer
Private AssRes As String
Public Sub About()
Dim MSG As String
MSG = "Encryptor ActiveX DLL" & Chr$(13)
MSG = MSG & "Version " & App.Major & "-" & App.Minor & "-" & App.Revision & Chr$(13)
MSG = MSG & "Written by Jack Hoxley" & Chr$(13)
MSG = MSG & "EMail: Jollyjeffers@hotmail.com"
MsgBox MSG
End Sub

Public Function Encrypt(StrEnc As String, IntVal As Integer) As String
If inval > 128 Then IntVal = 128
AssRes = ""
For i = 1 To Len(StrEnc)
Result = Mid$(StrEnc, i, 1)
Code = Asc(Result)
Code = Code + Int(IntVal)
If Code > 255 Then
Code = Code - 255
End If
AssRes = AssRes & Chr$(Code)
Next i
Encrypt = AssRes
End Function


Public Function Decrypt(StrEnc As String, IntVal As Integer) As String
If inval > 128 Then IntVal = 128
AssRes = ""
For i = 1 To Len(StrEnc)
Result = Mid$(StrEnc, i, 1)
Code = Asc(Result)
Code = Code - Int(IntVal)
If Code > 255 Then
Code = Code - 255
End If
AssRes = AssRes & Chr$(Code)
Next i
Decrypt = AssRes
End Function
