Attribute VB_Name = "ExcelTranslator"
Option Explicit

' =============================================================================
' Excel Dual-Engine Translator (Bing / Google) - No Shared Folder / No Admin Needed
' Works with the companion local HTTP endpoint (default http://127.0.0.1:3444).
' =============================================================================

Private Const APP_NAME As String = "ExcelDualTranslator"
Private Const DEFAULT_ENDPOINT As String = "http://127.0.0.1:3444"

Public Function GetActiveEngine() As String
    Dim eng As String
    eng = LCase$(Trim$(GetSetting(APP_NAME, "Config", "Engine", "bing")))
    If eng <> "bing" And eng <> "google" Then
        eng = "bing"
    End If
    GetActiveEngine = eng
End Function

Public Sub SetEngineBing()
    SaveSetting APP_NAME, "Config", "Engine", "bing"
    MsgBox "Translation engine switched to: BING", vbInformation, "Excel Translator"
End Sub

Public Sub SetEngineGoogle()
    SaveSetting APP_NAME, "Config", "Engine", "google"
    MsgBox "Translation engine switched to: GOOGLE", vbInformation, "Excel Translator"
End Sub

Public Sub SwitchTranslateEngine()
    Dim currentEngine As String
    Dim nextEngine As String
    currentEngine = GetActiveEngine()
    If currentEngine = "bing" Then
        nextEngine = "google"
    Else
        nextEngine = "bing"
    End If
    SaveSetting APP_NAME, "Config", "Engine", nextEngine
    MsgBox "Translation engine switched from " & UCase$(currentEngine) & " to " & UCase$(nextEngine), vbInformation, "Excel Translator"
End Sub

Public Sub ConfigureLanguages()
    Dim curFrom As String
    Dim curTo As String
    Dim newFrom As String
    Dim newTo As String

    curFrom = GetSetting(APP_NAME, "Config", "SourceLang", "auto")
    curTo = GetSetting(APP_NAME, "Config", "TargetLang", "zh-CN")

    newFrom = InputBox("Enter source language code (e.g. auto, en, zh-CN, ja):", "Source Language", curFrom)
    If StrPtr(newFrom) = 0 Then Exit Sub
    newTo = InputBox("Enter target language code (e.g. zh-CN, zh-TW, en, ja, ko):", "Target Language", curTo)
    If StrPtr(newTo) = 0 Then Exit Sub

    If Trim$(newFrom) <> "" Then SaveSetting APP_NAME, "Config", "SourceLang", Trim$(newFrom)
    If Trim$(newTo) <> "" Then SaveSetting APP_NAME, "Config", "TargetLang", Trim$(newTo)

    MsgBox "Saved: " & Trim$(newFrom) & " -> " & Trim$(newTo) & " (Engine: " & UCase$(GetActiveEngine()) & ")", vbInformation, "Excel Translator"
End Sub

' Translate active selection and fill immediately below (e.g. A1 -> A2)
Public Sub TranslateSelectionBelow()
    TranslateSelectionWithOffset True
End Sub

' Translate active selection and fill immediately to the right (e.g. A1 -> B1)
Public Sub TranslateSelectionRight()
    TranslateSelectionWithOffset False
End Sub

Private Sub TranslateSelectionWithOffset(ByVal fillBelow As Boolean)
    On Error GoTo ErrHandler
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select one or more cells first.", vbExclamation, "Excel Translator"
        Exit Sub
    End If

    Dim srcRange As Range
    Dim dstRange As Range
    Dim engine As String
    Dim fromLang As String
    Dim toLang As String
    Dim r As Long, c As Long
    Dim cellText As String
    Dim translated As String
    Dim countOk As Long

    Set srcRange = Selection
    If fillBelow Then
        Set dstRange = srcRange.Offset(srcRange.Rows.Count, 0)
    Else
        Set dstRange = srcRange.Offset(0, srcRange.Columns.Count)
    End If

    engine = GetActiveEngine()
    fromLang = GetSetting(APP_NAME, "Config", "SourceLang", "auto")
    toLang = GetSetting(APP_NAME, "Config", "TargetLang", "zh-CN")

    Application.StatusBar = "Translating via " & UCase$(engine) & "..."

    For r = 1 To srcRange.Rows.Count
        For c = 1 To srcRange.Columns.Count
            cellText = CStr(srcRange.Cells(r, c).Value)
            If Trim$(cellText) <> "" Then
                translated = CallTranslateApi(cellText, toLang, fromLang, engine)
                dstRange.Cells(r, c).Value = translated
                countOk = countOk + 1
            Else
                dstRange.Cells(r, c).Value = ""
            End If
        Next c
    Next r

    Application.StatusBar = False
    MsgBox "Done (" & UCase$(engine) & "): " & srcRange.Address(False, False) & " -> " & dstRange.Address(False, False) & " (" & countOk & " cells)", vbInformation, "Excel Translator"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "Translation error: " & Err.Description & vbCrLf & "Ensure local server is running (npm start).", vbCritical, "Excel Translator"
End Sub

' Custom Worksheet Function: =XLTRANS(A1, "zh-CN", "auto", "bing")
Public Function XLTRANS(ByVal text As Variant, Optional ByVal targetLang As String = "zh-CN", Optional ByVal sourceLang As String = "auto", Optional ByVal engine As String = "") As Variant
    On Error GoTo ErrFn
    Dim rawText As String
    rawText = CStr(text)
    If Trim$(rawText) = "" Then
        XLTRANS = ""
        Exit Function
    End If

    Dim activeEngine As String
    If Trim$(engine) = "" Then
        activeEngine = GetActiveEngine()
    Else
        activeEngine = LCase$(Trim$(engine))
    End If

    XLTRANS = CallTranslateApi(rawText, targetLang, sourceLang, activeEngine)
    Exit Function

ErrFn:
    XLTRANS = "#TRANS_ERR: " & Err.Description
End Function

Private Function CallTranslateApi(ByVal text As String, ByVal targetLang As String, ByVal sourceLang As String, ByVal engine As String) As String
    Dim http As Object
    Dim url As String
    Dim payload As String
    Dim respText As String

    url = GetSetting(APP_NAME, "Config", "Endpoint", DEFAULT_ENDPOINT) & "/api/translate"
    payload = "{""text"":""" & EscapeJsonString(text) & """,""engine"":""" & EscapeJsonString(engine) & """,""from"":""" & EscapeJsonString(sourceLang) & """,""to"":""" & EscapeJsonString(targetLang) & """}"

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send payload

    respText = http.responseText
    If http.Status < 200 Or http.Status >= 300 Then
        Dim errMsg As String
        errMsg = ExtractJsonStringField(respText, "error")
        If errMsg = "" Then errMsg = "HTTP " & CStr(http.Status)
        Err.Raise vbObjectError + 1001, "CallTranslateApi", errMsg
    End If

    Dim okFlag As String
    okFlag = ExtractJsonBoolField(respText, "ok")
    If okFlag = "false" Then
        Err.Raise vbObjectError + 1002, "CallTranslateApi", ExtractJsonStringField(respText, "error")
    End If

    CallTranslateApi = ExtractJsonStringField(respText, "translatedText")
End Function

Private Function EscapeJsonString(ByVal s As String) As String
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim sb As String

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "\"
                sb = sb & "\\"
            Case """"
                sb = sb & "\"""
            Case vbCr
                sb = sb & "\r"
            Case vbLf
                sb = sb & "\n"
            Case vbTab
                sb = sb & "\t"
            Case Else
                code = AscW(ch) And &HFFFF&
                If code < 32 Then
                    sb = sb & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    sb = sb & ch
                End If
        End Select
    Next i
    EscapeJsonString = sb
End Function

Private Function ExtractJsonBoolField(ByVal json As String, ByVal fieldName As String) As String
    Dim token As String
    Dim pos As Long
    token = """" & fieldName & """"
    pos = InStr(1, json, token, vbBinaryCompare)
    If pos = 0 Then Exit Function
    pos = InStr(pos + Len(token), json, ":", vbBinaryCompare)
    If pos = 0 Then Exit Function
    Dim rest As String
    rest = LTrim$(Mid$(json, pos + 1))
    If Left$(rest, 4) = "true" Then ExtractJsonBoolField = "true"
    If Left$(rest, 5) = "false" Then ExtractJsonBoolField = "false"
End Function

Private Function ExtractJsonStringField(ByVal json As String, ByVal fieldName As String) As String
    Dim token As String
    Dim pos As Long
    Dim i As Long
    Dim ch As String
    Dim nextCh As String
    Dim result As String
    Dim hexCode As String

    token = """" & fieldName & """"
    pos = InStr(1, json, token, vbBinaryCompare)
    If pos = 0 Then Exit Function

    pos = InStr(pos + Len(token), json, ":", vbBinaryCompare)
    If pos = 0 Then Exit Function

    pos = InStr(pos + 1, json, """", vbBinaryCompare)
    If pos = 0 Then Exit Function

    i = pos + 1
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        If ch = """" Then
            Exit Do
        ElseIf ch = "\" Then
            i = i + 1
            If i > Len(json) Then Exit Do
            nextCh = Mid$(json, i, 1)
            Select Case nextCh
                Case """": result = result & """"
                Case "\": result = result & "\"
                Case "/": result = result & "/"
                Case "n": result = result & vbLf
                Case "r": result = result & vbCr
                Case "t": result = result & vbTab
                Case "b": result = result & Chr$(8)
                Case "f": result = result & Chr$(12)
                Case "u"
                    If i + 4 <= Len(json) Then
                        hexCode = Mid$(json, i + 1, 4)
                        result = result & ChrW$(CLng("&H" & hexCode))
                        i = i + 4
                    End If
                Case Else
                    result = result & nextCh
            End Select
        Else
            result = result & ch
        End If
        i = i + 1
    Loop

    ExtractJsonStringField = result
End Function
