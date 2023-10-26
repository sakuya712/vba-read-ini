VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Lib")
Option Explicit

'===============================================================================
'''<classSummary> Read ini file (readonly)
'''<updateDate> 2023/10/25
'''<version> Ver.100
'===============================================================================

'Character code enum
Enum CharacterCodeEnum
    cceUtf8 = 65001
    cceShift_jis = 932
    cceAscii = 1252
End Enum

'Newline type enum
Enum NewlineCharacterEnum
    nceCRLF = 1310
    nceLF = 10
End Enum

Private Dic As Object

'''<summary>Constructor(call first)</summary>
'''<param name="FilePath">ini file full path</param>
'''<param name="CharacterCode">Character code</param>
'''<param name="NewlineCharacter">Newline character</param>
'''<returns>Me</returns>
'''<remarks>If the file is not found, returns "Nothing".</remarks>
Function Init(ByVal FilePath As String, Optional ByVal CharacterCode As CharacterCodeEnum = CharacterCodeEnum.cceUtf8, _
                Optional ByVal NewlineCharacter As NewlineCharacterEnum = NewlineCharacterEnum.nceCRLF) As Ini
Attribute Init.VB_Description = "Constructor(call first)"
    With CreateObject("Scripting.FileSystemObject")
        If Not (.FileExists(FilePath)) Then
            Set Init = Nothing
            Exit Function
        End If
    End With
    'CharacterCode string
    Dim Charset As String
    Select Case CharacterCode
        Case CharacterCodeEnum.cceUtf8
            Charset = "UTF-8"
        Case CharacterCodeEnum.cceShift_jis
            Charset = "Shift_JIS"
        Case CharacterCodeEnum.cceAscii
            Charset = "ascii"
        Case Else
            Charset = "UTF-8"
    End Select
    'Set newline type
    Dim NewlineChar As Variant
    Select Case NewlineCharacter
        Case NewlineCharacterEnum.nceCRLF
            NewlineChar = Constants.vbCrLf
        Case NewlineCharacterEnum.nceLF
            NewlineChar = Constants.vbLf
    End Select
    'read file
    Dim Data As Variant
    With CreateObject("ADODB.Stream")
        .Charset = Charset
        .Open
        .LoadFromFile FilePath
        Data = .Readtext
    End With
    Dim Lines As Variant
    Lines = VBA.Split(Data, NewlineChar)
    Dim Section As String
    Set Dic = CreateObject("Scripting.Dictionary")
    On Error Resume Next 'If keys (section & parameter) Duplicate, error.
    Dim i As Long
    For i = LBound(Lines) To UBound(Lines)
        If Len(Lines(i)) = 0 Then GoTo Continue
        Select Case Left(Lines(i), 1)
            Case ";", "#", "/"
                'Treated as comment out
                GoTo Continue
            Case "["
                'Treated as section
                Dim RightIndex As Long
                RightIndex = InStr(Lines(i), "]")
                If RightIndex > 0 Then Section = Mid(Lines(i), 2, RightIndex - 2)
            Case Else
                'Treated as data
                If InStr(Lines(i), "=") > 0 Then
                    Dim SplitArray As Variant
                    SplitArray = Split(Lines(i), "=")
                    Dic.Add Section & SplitArray(0), SplitArray(1)
                End If
        End Select
Continue:
    Next
    Set Init = Me
End Function

'''<summary>Get value</summary>
'''<param name="Section">section</param>
'''<param name="Parameter">parameter</param>
'''<returns>value</returns>
'''<remarks>If not match found, returns "".</remarks>
Function GetValue(ByVal Section As String, ByVal Parameter As String) As String
Attribute GetValue.VB_Description = "Get value"
    On Error GoTo KeyNotFoundException
    GetValue = Dic(Section & Parameter)
    Exit Function
KeyNotFoundException:
    GetValue = ""
End Function
