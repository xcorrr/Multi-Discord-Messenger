Dim http, url, message, payload
' Ask user for webhook URL
url = InputBox("Webhook URL:", "Webhook Messenger - xcorr")

If url = "" Then
    MsgBox "Error: No URL provided.", vbOkOnly, "Webhook Messenger."
    WScript.Quit
End If
' Ask user for message
message = InputBox("Enter the message to send:", "Webhook Messenger - xcorr")

If message = "" Then
    MsgBox "Error: No message entered.", vbOkOnly, "Webhook Messenger."
    WScript.Quit
End If
' Create JSON payload
payload = "{""content"": """ & Replace(message, """", "\""") & """}"
' Send the request
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "POST", url, False
http.setRequestHeader "Content-Type", "application/json"
http.Send payload
' Confirm send
If http.Status = 204 Then
    MsgBox "Sucess: Message sent!", vbOKOnly, "Webhook Messenger."
Else
    MsgBox "Error: Failed sending message. Status: " & http.Status & vbCrLf & http.responseText 
End If