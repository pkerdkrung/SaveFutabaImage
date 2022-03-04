Sub Download_2chan_Image()

    Dim fso As Object
    Dim IE As Object
    Dim URL As String
    Dim SavePath As String
    Dim ImgCount As Integer
    Dim atags As Object
    Dim atag As Object
    Dim ImageUrl As String
    Dim Imagename As String
    
    Dim objElement As Object
    Dim objCollection As Object
    
    'Create FSO Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
    
    'Set IE.Visible = True to make IE visible, or False for IE to run in the background
    IE.Visible = False
    
    'Ask the user to input 2chan thread Url. If the user enter nothing, exit the sub
    URL = Application.InputBox(prompt:="Please enter Url", Type:=2)
    If URL = "" Then
        MsgBox "No Url. Program will terminate."
        Exit Sub
    End If
    
    'Pop up the folder-selection box to get the folder form the user:
    SavePath = GetFolder()
    ' If the user didn't select anything, you can't save, so tell them so:
    If SavePath = "" Then
        MsgBox "No folder was selected. Program will terminate."
        Exit Sub
    End If

    'Minimize the Excel window while the code runs
    ActiveWindow.WindowState = xlMinimized

    'Navigate to URL
    IE.navigate URL
    
    ' Statusbar let's user know website is loading
    Application.StatusBar = URL & " is loading. Please wait..."
    
    ' Wait while IE loading...
    'IE ReadyState = 4 signifies the webpage has loaded (the first loop is set to avoid inadvertently skipping over the second loop)
    Do While IE.readyState = 4: DoEvents: Loop   'Do While
    Do Until IE.readyState = 4: DoEvents: Loop   'Do Until
    
    'Webpage Loaded
    Application.StatusBar = URL & " Loaded"

    'Loop through each html with TagName "a". Keep count for each image downloaded.
    Set atags = IE.document.getElementsByTagName("a")
    ImgCount = 0
    For Each atag In atags
      ImageUrl = atag.href
      Imagename = fso.GetFileName(ImageUrl)
      On Error Resume Next
      If Imagename Like "*.png" Or Imagename Like "*.jpg" Or Imagename Like "*.gif" Or Imagename Like "*.webm" Or Imagename Like "*.mp4" Then
        ImgCount = ImgCount + 1
        Application.StatusBar = "Downloading. Total " & ImgCount & " files downloaded..."
        Call DownloadFileFromURL(ImageUrl, SavePath, Imagename)
        On Error Resume Next
      End If
    Next atag
    
    'Unload IE
    Set IE = Nothing
    Set objElement = Nothing
    Set objCollection = Nothing
    
    'Maximize the Excel window and alert the user
    ActiveWindow.WindowState = xlMaximized
    MsgBox "Download Completed. Total " & ImgCount & " files downloaded."
    
End Sub

'A Sub to download a file from the internet given Url, SavePath, and Filename.
Sub DownloadFileFromURL(varImageUrl, varSavePath As String, varImagename As String)

     Dim FileUrl As String
     Dim objXmlHttpReq As Object
     Dim objStream As Object

     FileUrl = varImageUrl
     
     Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
     objXmlHttpReq.Open "GET", FileUrl, False, "username", "password"
     objXmlHttpReq.send

     If objXmlHttpReq.Status = 200 Then
          Set objStream = CreateObject("ADODB.Stream")
          objStream.Open
          objStream.Type = 1
          objStream.Write objXmlHttpReq.responseBody
          objStream.SaveToFile varSavePath & "\" & varImagename, 2
          objStream.Close
     End If

End Sub

'A function to call up the folder-select dialog
Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With

NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function
