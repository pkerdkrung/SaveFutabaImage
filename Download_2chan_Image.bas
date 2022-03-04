'I like to save anime screenshots from Futaba Channel to make memes later,
'so I made a macro that downloads all images and video (png, jpg, gif, webm, mp4) in a thread.
'The user needs to enter a URL and specify a destination folder.
Option Explicit
Sub Download_2chan_Image()
	Dim objFSO As Object
	Dim objIE As Object
	Dim strURL As String
	Dim strSavePath As String
	Dim intImgCount As Integer
	Dim objAtags As Object
	Dim objAtag As Object
	Dim strImageUrl As String
	Dim strImagename As String
	Dim objElement As Object
	Dim objCollection As Object
	
	'Create FSO Object
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	'Create InternetExplorer Object
  Set objIE = CreateObject("InternetExplorer.Application")
    
  'Set IE.Visible = True to make IE visible, or False for IE to run in the background
  objIE.Visible = False
    
  'Ask the user to input 2chan thread Url. If the user enter nothing, exit the sub
  strURL = ""
  strURL = Application.InputBox(prompt:="Please enter Url", Type:=2)
  If strURL = "" Then
  	MsgBox "No Url. Program will terminate."
    Exit Sub
  End If
    
  'Pop up the folder-selection box to get the folder form the user:
  strSavePath = GetFolder()
  
	' If the user didn't select anything, you can't save, so tell them so:
  If strSavePath = "" Then
  	MsgBox "No folder was selected. Program will terminate."
    Exit Sub
  End If

  'Minimize the Excel window while the code runs
  ActiveWindow.WindowState = xlMinimized

  'Navigate to URL
  objIE.navigate strURL
    
  ' Statusbar let's user know website is loading
  Application.StatusBar = strURL & " is loading. Please wait..."
    
  ' Wait while IE loading...
  'IE ReadyState = 4 signifies the webpage has loaded (the first loop is set to avoid inadvertently skipping over the second loop)
  Do While objIE.readyState = 4: DoEvents: Loop   'Do While
  Do Until objIE.readyState = 4: DoEvents: Loop   'Do Until
    
  'Webpage Loaded
  Application.StatusBar = strURL & " Loaded"

  'Loop through each html with TagName "a". Keep count for each image downloaded.
  Set objAtags = objIE.document.getElementsByTagName("a")
  intImgCount = 0
  For Each objAtag In objAtags
  	strImageUrl = objAtag.href
    strImagename = objFSO.GetFileName(ImageUrl)
    On Error Resume Next
    If strImagename Like "*.png" Or strImagename Like "*.jpg" Or strImagename Like "*.gif" Or strImagename Like "*.webm" Or strImagename Like "*.mp4" Then
    	intImgCount = intImgCount + 1
      Application.StatusBar = "Downloading. Total " & intImgCount & " files downloaded..."
      Call DownloadFileFromURL(strImageUrl, strSavePath, strImagename)
      On Error Resume Next
    End If
  Next objAtag
    
  'Unload IE
  Set objIE = Nothing
  Set objElement = Nothing
  Set objCollection = Nothing
    
  'Maximize the Excel window and alert the user
  ActiveWindow.WindowState = xlMaximized
  MsgBox "Download Completed. Total " & intImgCount & " files downloaded."
    
End Sub

'A helper routine to download a file from the internet given Url, SavePath, and Filename.
     Sub DownloadFileFromURL(varImageUrl As String, varSavePath As String, varImagename As String)

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
