Attribute VB_Name = "LoopIn"
Option Explicit

Private Function RegRead(key As String) As String
    On Error Resume Next
    RegRead = CreateObject("WScript.Shell").RegRead("HKCU\Software\LoopIn\" & key)
    If Err.Number <> 0 Then RegRead = ""
    Err.Clear
End Function

Private Sub RegWrite(key As String, value As String)
    CreateObject("WScript.Shell").RegWrite "HKCU\Software\LoopIn\" & key, value, "REG_SZ"
End Sub

Public Sub LoopIn_Send()
    Dim webhook As String
    webhook = RegRead("Webhook")

    If webhook = "" Then
        MsgBox "No webhook URL saved. Run LoopIn_Setup first.", vbExclamation, "LoopIn"
        LoopIn_Setup
        Exit Sub
    End If

    Dim channel As String
    channel = RegRead("Channel")
    If channel = "" Then channel = "#general"

    Dim choice As String
    choice = InputBox("Choose a template (type 1, 2, or 3) or leave blank to write your own:" & Chr(10) & Chr(10) & _
        "1 - Dashboard Update" & Chr(10) & _
        "2 - Sprint Summary" & Chr(10) & _
        "3 - Gone Live" & Chr(10), "LoopIn - Send to " & channel)

    Dim msg As String
    Select Case Trim(choice)
        Case "1"
            msg = ":bar_chart: *Salesforce Projects Dashboard Updated*" & Chr(10) & Chr(10) & _
                  "The latest sprint data is now live. Check current sprint capacity, project states, and upcoming releases."
        Case "2"
            msg = ":runner: *Sprint Update - Quilt Software*" & Chr(10) & Chr(10) & _
                  "Sprint planning has been updated. Review what's in flight, what's in prep, and what's coming up next."
        Case "3"
            msg = ":tada: *Projects Gone Live*" & Chr(10) & Chr(10) & _
                  "New projects have been marked as Gone Live! Check the dashboard for the full release history."
        Case ""
            Exit Sub
        Case Else
            msg = InputBox("Type your message:", "LoopIn - Send to " & channel)
            If Trim(msg) = "" Then Exit Sub
    End Select

    If MsgBox("Send to " & channel & "?" & Chr(10) & Chr(10) & msg, vbYesNo + vbQuestion, "LoopIn") = vbNo Then Exit Sub

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", webhook, False
    http.SetRequestHeader "Content-Type", "application/json"

    Dim escaped As String
    escaped = Replace(msg, "\", "\\")
    escaped = Replace(escaped, """", "\""")
    escaped = Replace(escaped, Chr(10), "\n")

    On Error GoTo SendErr
    http.Send "{""text"": """ & escaped & """}"

    If http.Status = 200 Then
        MsgBox "Sent to " & channel & "!", vbInformation, "LoopIn"
    Else
        MsgBox "Slack returned an error. Check your webhook URL.", vbExclamation, "LoopIn"
    End If
    Exit Sub

SendErr:
    MsgBox "Failed to connect. Check your internet connection and webhook URL.", vbCritical, "LoopIn"
End Sub

Public Sub LoopIn_Setup()
    Dim webhook As String
    webhook = InputBox("Paste your Slack Incoming Webhook URL:", "LoopIn Setup", RegRead("Webhook"))
    If Trim(webhook) = "" Then Exit Sub
    RegWrite "Webhook", Trim(webhook)

    Dim channel As String
    channel = InputBox("Channel name (e.g. #maryworkflowtests):", "LoopIn Setup", RegRead("Channel"))
    If Trim(channel) = "" Then channel = "#general"
    RegWrite "Channel", Trim(channel)

    MsgBox "LoopIn is ready! Run LoopIn_Send to send messages.", vbInformation, "LoopIn"
End Sub
