Option Explicit

Dim objWMIService, objItem, colItems
Dim objHTTP, strPublicIP, strLocalIP, strURL, strResult

' Getting the local IP address
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")

For Each objItem in colItems
    If Not IsNull(objItem.IPAddress) Then
        strLocalIP = Join(objItem.IPAddress, ", ")
        Exit For
    End If
Next

' URL of the service to retrieve the public IP address
strURL = "https://api64.ipify.org?format=json"

' Creating the XMLHTTP object
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

' Sending a GET request to the service
Dim startTime
startTime = Timer
objHTTP.open "GET", strURL, False
objHTTP.send
Dim responseTime
responseTime = Int((Timer - startTime) * 1000) ' Conversion to milliseconds without decimals

' Retrieving the response
strResult = objHTTP.responseText

' Constructing the strings to display
Dim displayString
displayString = "Local IP: " & strLocalIP & vbCrLf

If InStr(strResult, """ip"":""") > 0 Then
    Dim startPos, endPos
    startPos = InStr(strResult, """ip"":""") + Len("""ip"":""")
    endPos = InStr(startPos, strResult, """") - 1
    strPublicIP = Mid(strResult, startPos, endPos - startPos + 1)
    displayString = displayString & "Public IP: " & strPublicIP & vbCrLf
    displayString = displayString & "Latency: " & responseTime & " ms"
    
    ' Copying the public IP address to the clipboard
    CopyToClipboard strPublicIP
    
    ' Skipping two lines and indicating that the public IP is copied to the clipboard
    displayString = displayString & vbCrLf & vbCrLf
    displayString = displayString & "The public IP address has been copied to the clipboard."
Else
    displayString = displayString & "Unable to retrieve the public IP address."
End If

' Displaying the information in the same window
MsgBox displayString, vbInformation, "Get_IP by Azukiov"

' Releasing the HTTP object
Set objHTTP = Nothing

' Function to copy to the clipboard
Sub CopyToClipboard(ByVal strText)
    Dim objShell : Set objShell = CreateObject("WScript.Shell")
    objShell.Run "cmd /c echo " & strText & " | clip", 0, True
    Set objShell = Nothing
End Sub
