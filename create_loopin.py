"""Creates LoopIn.xlam — Excel Add-In that sends Slack messages."""

import win32com.client as win32
import os

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "LoopIn.xlam")

VBA_CODE = '''
Option Explicit

' Store settings in Windows registry
Private Function RegRead(key As String) As String
    On Error Resume Next
    RegRead = CreateObject("WScript.Shell").RegRead("HKCU\\Software\\LoopIn\\" & key)
    If Err.Number <> 0 Then RegRead = ""
    Err.Clear
End Function

Private Sub RegWrite(key As String, value As String)
    CreateObject("WScript.Shell").RegWrite "HKCU\\Software\\LoopIn\\" & key, value, "REG_SZ"
End Sub

' ── Main: Send to Slack ─────────────────────────────────────────────────────
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

    ' Template choice
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

    ' Confirm before sending
    If MsgBox("Send to " & channel & "?" & Chr(10) & Chr(10) & msg, vbYesNo + vbQuestion, "LoopIn") = vbNo Then Exit Sub

    ' Send to Slack
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", webhook, False
    http.SetRequestHeader "Content-Type", "application/json"

    Dim payload As String
    Dim escaped As String
    escaped = Replace(msg, "\", "\\")
    escaped = Replace(escaped, """", "\""")
    escaped = Replace(escaped, Chr(10), "\n")
    payload = "{""text"": """ & escaped & """}"

    On Error GoTo SendErr
    http.Send payload

    If http.Status = 200 Then
        MsgBox "Sent to " & channel & "!", vbInformation, "LoopIn"
    Else
        MsgBox "Slack returned an error. Check your webhook URL.", vbExclamation, "LoopIn"
    End If
    Exit Sub

SendErr:
    MsgBox "Failed to connect. Check your internet connection and webhook URL.", vbCritical, "LoopIn"
End Sub

' ── Setup: Save webhook URL and channel ─────────────────────────────────────
Public Sub LoopIn_Setup()
    Dim webhook As String
    webhook = InputBox("Paste your Slack Incoming Webhook URL:", "LoopIn Setup", RegRead("Webhook"))
    If Trim(webhook) = "" Then Exit Sub
    RegWrite "Webhook", Trim(webhook)

    Dim channel As String
    channel = InputBox("Channel name (e.g. #general):", "LoopIn Setup", RegRead("Channel"))
    If Trim(channel) = "" Then channel = "#general"
    RegWrite "Channel", Trim(channel)

    MsgBox "LoopIn is ready! Use LoopIn_Send to send messages.", vbInformation, "LoopIn"
End Sub
'''

def build():
    print("Opening Excel...")
    xl = win32.Dispatch("Excel.Application")
    xl.DisplayAlerts = False

    wb = xl.Workbooks.Add()

    try:
        mod = wb.VBProject.VBComponents.Add(1)  # standard module
        mod.Name = "LoopIn"
        mod.CodeModule.AddFromString(VBA_CODE)
        print("VBA module added.")

        if os.path.exists(OUT):
            os.remove(OUT)
        wb.SaveAs(OUT, 55)  # 55 = xlOpenXMLAddIn (.xlam)
        print(f"\nSuccess! Saved to:\n{OUT}")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        wb.Close(False)
        xl.Quit()

if __name__ == "__main__":
    build()
