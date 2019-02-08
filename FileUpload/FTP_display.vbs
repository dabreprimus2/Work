dim objConnection
Dim objRecordSet
dim mbr
dim rename
'Database Connection
Sub ConnectToDatabase()
    dim strConnectionString
    dim strServerName
    dim strDatabase 
    dim strUserID 
    dim strPassword

	strServerName = "Printers.Hi-Techhealth.com"
	strDatabase = "S10BB55T"
	strUserID = App.GetValue("varUserName")
	strPassword = App.GetValue("varPassword")

    'Create a new instance of the Connection Object
    Set objConnection = CreateObject("ADODB.Connection")
	
	strConnectionString = "Provider=IBMDA400;Data Source=" & strServerName & ";User Id=" & strUserID & ";Password=" & strPassword
	'strTable  = "SECNAM"
	'Connect to the database
	objConnection.Open strConnectionString
End Sub


 
Function GetFileDlgEx(sIniDir,sFilter,sTitle) 
	Set oDlg = CreateObject("WScript.Shell").Exec("mshta.exe ""about:<object id=d classid=clsid:3050f4e1-98b5-11cf-bb82-00aa00bdce0b></object><script>moveTo(0,-9999);eval(new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(0).Read("&Len(sIniDir)+Len(sFilter)+Len(sTitle)+41&"));function window.onload(){var p=/[^\0]*/;new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).Write(p.exec(d.object.openfiledlg(iniDir,null,filter,title)));close();}</script><hta:application showintaskbar=no />""") 
	oDlg.StdIn.Write "var iniDir='" & sIniDir & "';var filter='" & sFilter & "';var title='" & sTitle & "';" 
	GetFileDlgEx = oDlg.StdOut.ReadAll 
End Function

sub newtest
	NewItemIndex = CInt(App.ActiveForm.Controls.fileTypeLst.ListIndex)
	set fso = CreateObject("Scripting.FileSystemObject")
	CurrentDirectory = fso.GetAbsolutePathName(".")
	sIniDir = CurrentDirectory 
	sFilter = App.ActiveForm.Controls.fileTypeLst.List (NewItemIndex)
	sTitle = "Upload File" 
	MyFile = GetFileDlgEx(Replace(sIniDir,"\","\\"),sFilter,sTitle) 

	If MyFile = "" Then
		MsgBox "Operation canceled", vbcritical
	Else
		dim sfileName
	'gets the file name from the local path
		sfileName= mid(MyFile, InstrRev(MyFile,"\")+1,len(MyFile))
		App.SetValue "lblSetFilename", sfileName 
		App.SetValue "setDir",MyFile
		'MsgBox MyFile, vbinformation
	End If
	
end sub
Sub FileUpload
	dim check
	GrpID = App.GetValue("Grptxt")
	DivID = App.GetValue("Divtxt")
	EmpID = App.GetValue("Emptxt")
	dim path
	path = App.GetValue("setDir")
	GrpID = Trim(GrpID)
	DivID = Trim(DivID)
	EmpID = Trim(EmpID)
	If GrpID <> "" and DivID ="" and EmpID ="" then
		
		check = App.ActiveForm.lblSetFilename.Caption
		if check <> "No file selected" then
		GrpFile(path)
		else
		MsgBox "Select a file to upload", vbcritical	
		end if
	End If
	
	If GrpID <> "" and DivID <> "" and EmpID="" then
		check = App.ActiveForm.lblSetFilename.Caption
		if check <> "No file selected" then
		DivFile(path)
		else
		MsgBox "Select a file to upload", vbcritical	
		end if

	End If
	
	If GrpID <> "" and EmpID <> "" then
	
		check = App.ActiveForm.lblSetFilename.Caption
		if check <> "No file selected" then
		EmpFile(path)
		else
		MsgBox "Select a file to upload", vbcritical	
		end if
	End If
End Sub

sub test
NewItemIndex = CInt(App.ActiveForm.Controls.cmbType.ListIndex)
	msgbox NewItemIndex  
end sub


Function GrpFile(MyFile)
	
	'App.ActiveForm.Controls.cmbType.List (NewItemIndex)
	GrpID= App.GetValue ("Grptxt")
	Set objShell = CreateObject("WScript.Shell")
	currentFolder = objShell.CurrentDirectory
	dim srclst
	srclst = MyFile
	dim sfileExt
	sfileExt = srclst
	'gets the file extension
	sfileExt= mid(sfileExt, InstrRev(sfileExt,"\")+1,len(sfileExt))
	sfileExt= mid(sfileExt, InstrRev(sfileExt,".")+1,len(sfileExt))
	'msgbox sfileExt
	Dim objFileSystemOjbect
	Set objFileSystemOjbect = CreateObject("Scripting.FileSystemObject")
	Set appShell = CreateObject("Shell.Application")

	destinationFolderPath = objShell.ExpandEnvironmentStrings("%temp%") & "\MyApp_TMP\"
	'Check if folder exists
	Dim newfolder
	If Not objFileSystemOjbect.FolderExists(destinationFolderPath) Then
    'if not then create one
    Set newfolder = objFileSystemOjbect.CreateFolder(destinationFolderPath)
	Else
    'clear all existing files under the destination folder
    objFileSystemOjbect.DeleteFile(destinationFolderPath & "\*"), true  ' true - delete read only files
	End If
	rename = Trim(App.GetValue("txtRename"))&ConvertDate(FormatDateTime(Now()))&"."&sfileExt
	dim desc
	desc = App.GetValue("txtRename")&"."&sfileExt
	objFileSystemOjbect.CopyFile MyFile, destinationFolderPath &rename,true
	FTPUpload(destinationFolderPath&rename)	
	QueryDatabase
	sRemotePath = "/web/"&mbr&"/attachments" 'pass a member
	Set objRecordSet = CreateObject("ADODB.Recordset")
	'Define your SQL statement
	objRecordSet.Source = "insert into cblib.imgmisc values ('"&sRemotePath&"/"&rename&"','"&GrpID&"','"&FormatDateTime(Now(),2)&"','GROUP','"&GrpID&"GF','"&desc&"','"&GrpID&"','','')"
	Set objRecordSet.ActiveConnection = objConnection
	objRecordSet.Open
	Set objRecordSet = Nothing 
end function

Function DivFile(MyFile)
	NewItemIndex = CInt(App.ActiveForm.Controls.cmbType.ListIndex)
	App.ActiveForm.Controls.cmbType.List (NewItemIndex)
	GrpID= App.GetValue ("Grptxt")
	Set objShell = CreateObject("WScript.Shell")
	currentFolder = objShell.CurrentDirectory
	dim srclst
	srclst = MyFile
	dim sfileExt
	sfileExt = srclst
	'gets the file extension
	sfileExt= mid(sfileExt, InstrRev(sfileExt,"\")+1,len(sfileExt))
	sfileExt= mid(sfileExt, InstrRev(sfileExt,".")+1,len(sfileExt))
	'msgbox sfileExt
	Dim objFileSystemOjbect
	Set objFileSystemOjbect = CreateObject("Scripting.FileSystemObject")
	Set appShell = CreateObject("Shell.Application")

	destinationFolderPath = objShell.ExpandEnvironmentStrings("%temp%") & "\MyApp_TMP\"
	'Check if folder exists
	Dim newfolder
	If Not objFileSystemOjbect.FolderExists(destinationFolderPath) Then
    'if not then create one
    Set newfolder = objFileSystemOjbect.CreateFolder(destinationFolderPath)
	Else
    'clear all existing files under the destination folder
    objFileSystemOjbect.DeleteFile(destinationFolderPath & "\*"), true  ' true - delete read only files
	End If
	rename = Trim(App.GetValue("txtRename"))&ConvertDate(FormatDateTime(Now()))&"."&sfileExt
	dim desc
	desc = App.GetValue("txtRename")&"."&sfileExt
	objFileSystemOjbect.CopyFile MyFile, destinationFolderPath &rename,true
	FTPUpload(destinationFolderPath&rename)	
	QueryDatabase
	sRemotePath = "/web/"&mbr&"/attachments" 'pass a member
	Set objRecordSet = CreateObject("ADODB.Recordset")
	'Define your SQL statement
	objRecordSet.Source = "insert into cblib.imgmisc values ('"&sRemotePath&"/"&rename&"','"&GrpID&"','"&FormatDateTime(Now(),2)&"','DIVISION','"&GrpID&"DF','"&desc&"','"&GrpID&"',)"
	Set objRecordSet.ActiveConnection = objConnection
	objRecordSet.Open
	Set objRecordSet = Nothing 
end function

Function EmpFile(MyFile)
	NewItemIndex = CInt(App.ActiveForm.Controls.cmbType.ListIndex)
	App.ActiveForm.Controls.cmbType.List (NewItemIndex)
	GrpID= App.GetValue ("Grptxt")
	Set objShell = CreateObject("WScript.Shell")
	currentFolder = objShell.CurrentDirectory
	dim srclst
	srclst = MyFile
	dim sfileExt
	sfileExt = srclst
	'gets the file extension
	sfileExt= mid(sfileExt, InstrRev(sfileExt,"\")+1,len(sfileExt))
	sfileExt= mid(sfileExt, InstrRev(sfileExt,".")+1,len(sfileExt))
	'msgbox sfileExt
	Dim objFileSystemOjbect
	Set objFileSystemOjbect = CreateObject("Scripting.FileSystemObject")
	Set appShell = CreateObject("Shell.Application")

	destinationFolderPath = objShell.ExpandEnvironmentStrings("%temp%") & "\MyApp_TMP\"
	'Check if folder exists
	Dim newfolder
	If Not objFileSystemOjbect.FolderExists(destinationFolderPath) Then
    'if not then create one
    Set newfolder = objFileSystemOjbect.CreateFolder(destinationFolderPath)
	Else
    'clear all existing files under the destination folder
    objFileSystemOjbect.DeleteFile(destinationFolderPath & "\*"), true  ' true - delete read only files
	End If
	rename = Trim(App.GetValue("txtRename"))&ConvertDate(FormatDateTime(Now()))&"."&sfileExt
	dim desc
	desc = App.GetValue("txtRename")&"."&sfileExt
	objFileSystemOjbect.CopyFile MyFile, destinationFolderPath &rename,true
	FTPUpload(destinationFolderPath&rename)	
	QueryDatabase
	sRemotePath = "/web/"&mbr&"/attachments" 'pass a member
	Set objRecordSet = CreateObject("ADODB.Recordset")
	'Define your SQL statement
	objRecordSet.Source = "insert into cblib.imgmisc values ('"&sRemotePath&"/"&rename&"','"&GrpID&"','"&FormatDateTime(Now(),2)&"','EMPLOYEE','"&GrpID&"EF','"&desc&"','"&GrpID&"')"
	Set objRecordSet.ActiveConnection = objConnection
	objRecordSet.Open
	Set objRecordSet = Nothing 
end function

Sub getName()
	dim FrName, FrCaption 
	FrName = App.ActiveForm.Name
	FrCaption = App.ActiveForm.Caption
	If FrName = "GRPMNT" then
		GrpID = App.ActiveForm.lblGrpID.Caption
		App.ShowForms=False
		App.OpenForm "FileChooser"
		App.SetValue "Grptxt", GrpID
		App.SetValue "txtRename", Trim(GrpID)
		NewItem = Array("PDF Files (*.pdf)|*.pdf|","Document Files (*.docx)|*.docx|","Image files (*.png)|*.png|","All Files (*.*)|*.*")
         b = ubound(NewItem)
             For i = 0 to b
		NewItemIndex = CInt(App.ActiveForm.Controls.fileTypeLst.ListCount)
		AddNewSuccess = App.ActiveForm.Controls.fileTypeLst.AddItem (NewItem(i),NewItemIndex)
		'msgbox"New Item: " &NewItem(i)
		'msgbox"NewItemIndex: " &NewItemIndex
		Next
		'cmbType
		NewItem =Array ("Both","Internal","Web")

             For i = 0 to 2
		
		NewItemIndex = CInt(App.ActiveForm.Controls.cmbType.ListCount)
		AddNewSuccess = App.ActiveForm.Controls.cmbType.AddItem (NewItem(i),NewItemIndex)
		'msgbox"New Item: " &NewItem(i)
		'msgbox"NewItemIndex: " &NewItemIndex
		Next
	App.ShowForms=True
	End If
	If FrName = "DIVMNT" then
		DivID = App.ActiveForm.lblDivID.Caption
		GrpID = App.ActiveForm.lblGrpID.Caption
		App.ShowForms=False
		App.OpenForm "FileChooser"
		App.SetValue "Grptxt", GrpID
		App.SetValue "Divtxt", DivID
		App.SetValue "txtRename", Trim(GrpID)&Trim(DivID)
		NewItem = Array("PDF Files (*.pdf)|*.pdf|","Document Files (*.docx)|*.docx|","Image files (*.png)|*.png|","All Files (*.*)|*.*")
         b = ubound(NewItem)
             For i = 0 to b
		NewItemIndex = CInt(App.ActiveForm.Controls.fileTypeLst.ListCount)
		AddNewSuccess = App.ActiveForm.Controls.fileTypeLst.AddItem (NewItem(i),NewItemIndex)
		Next
		'cmbType
		NewItem =Array ("Both","Internal","Web")

             For i = 0 to 2
		
		NewItemIndex = CInt(App.ActiveForm.Controls.cmbType.ListCount)
		AddNewSuccess = App.ActiveForm.Controls.cmbType.AddItem (NewItem(i),NewItemIndex)
		Next
		App.ShowForms=True
	End If

	If FrName = "EMPMNT" then
		'DivID = App.ActiveForm.lblDivID.Caption
		GrpID = App.ActiveForm.lblGrpID.Caption
		EmpID = App.ActiveForm.lblEmpID.Caption
		App.ShowForms=False
		App.OpenForm "FileChooser"
		App.SetValue "Grptxt", GrpID
		'App.SetValue "Divtxt", DivID
		App.SetValue "Emptxt", EmpID
		App.SetValue "txtRename", Trim(GrpID)&Trim(EmpID)
		NewItem = Array("PDF Files (*.pdf)|*.pdf|","Document Files (*.docx)|*.docx|","Image files (*.png)|*.png|","All Files (*.*)|*.*")
         b = ubound(NewItem)
             For i = 0 to b
		NewItemIndex = CInt(App.ActiveForm.Controls.fileTypeLst.ListCount)
		AddNewSuccess = App.ActiveForm.Controls.fileTypeLst.AddItem (NewItem(i),NewItemIndex)
		
		Next
		'cmbType
		NewItem =Array ("Both","Internal","Web")
             For i = 0 to 2
		
		NewItemIndex = CInt(App.ActiveForm.Controls.cmbType.ListCount)
		AddNewSuccess = App.ActiveForm.Controls.cmbType.AddItem (NewItem(i),NewItemIndex)
		Next
		App.ShowForms=True
	End If

End Sub

Function FTPUpload(sLocalFile)
  Const OpenAsDefault = -2
  Const FailIfNotExist = 0
  Const ForReading = 1
  Const ForWriting = 2
  'connect to dababase and get menber
  'msgbox sLocalFile
  ConnectToDatabase
  QueryDatabase
	mbr =Trim(mbr)
	Set oFTPScriptFSO = CreateObject("Scripting.FileSystemObject")
	Set oFTPScriptShell = CreateObject("WScript.Shell")
	sSite ="209.191.33.19" 'ftp Address
	'app login credentials
	sUsername =	App.GetValue("varUserName") 
	sPassword =	App.GetValue("varPassword")
	sRemotePath = "/web/"&mbr&"/attachments" 'pass a member
	sRemotePath = Trim(sRemotePath)
	sLocalFile = Trim(sLocalFile) 
	dim sfileName
	'gets the file name from the local path
	sfileName= mid(sLocalFile, InstrRev(sLocalFile,"\")+1,len(sLocalFile))
	'msgbox sfileName
  '----------Path Checks---------
  'Here we willcheck the path, if it contains
  'spaces then we need to add quotes to ensure
  'it parses correctly.
  If InStr(sRemotePath, " ") > 0 Then
    If Left(sRemotePath, 1) <> """" And Right(sRemotePath, 1) <> """" Then
      sRemotePath = """" & sRemotePath & """"
    End If
  End If
  
  If InStr(sLocalFile, " ") > 0 Then
    If Left(sLocalFile, 1) <> """" And Right(sLocalFile, 1) <> """" Then
      sLocalFile = """" & sLocalFile & """"
    End If
  End If

  If Len(sRemotePath) = 0 Then

    sRemotePath = "\"
  End If

  If InStr(sLocalFile, "*") Then
    If InStr(sLocalFile, " ") Then
      FTPUpload = "Error: Wildcard uploads do not work if the path contains a " & _
      "space." & vbCRLF
      FTPUpload = FTPUpload & "This is a limitation of the Microsoft FTP client."
		msgbox FTPUpload
      Exit Function
    End If
  ElseIf Len(sLocalFile) = 0 Or Not oFTPScriptFSO.FileExists(sLocalFile) Then
    'nothing to upload
    FTPUpload = "Error: File Not Found."
	msgbox FTPUpload
    Exit Function
  End If

  
  'build input file for ftp command
  sFTPScript = sFTPScript & "USER " & sUsername & vbCRLF
  sFTPScript = sFTPScript & sPassword & vbCRLF
  sFTPScript = sFTPScript & "cd " & sRemotePath & vbCRLF
  sFTPScript = sFTPScript & "binary" & vbCRLF
  sFTPScript = sFTPScript & "prompt n" & vbCRLF
  sFTPScript = sFTPScript & "put " & sLocalFile & vbCRLF
  sFTPScript = sFTPScript & "quit" & vbCRLF & "quit" & vbCRLF & "quit" & vbCRLF


  sFTPTemp = oFTPScriptShell.ExpandEnvironmentStrings("%TEMP%")
  sFTPTempFile = sFTPTemp & "\" & oFTPScriptFSO.GetTempName
  sFTPResults = sFTPTemp & "\" & oFTPScriptFSO.GetTempName

  Set fFTPScript = oFTPScriptFSO.CreateTextFile(sFTPTempFile, True)
  fFTPScript.WriteLine(sFTPScript)
  fFTPScript.Close
  Set fFTPScript = Nothing  

  oFTPScriptShell.Run "%comspec% /c FTP -n -s:" & sFTPTempFile & " " & sSite & _
  " > " & sFTPResults, 0, TRUE
  
  
  'Check results of transfer.
  Set fFTPResults = oFTPScriptFSO.OpenTextFile(sFTPResults, ForReading, _
  FailIfNotExist, OpenAsDefault)
  sResults = fFTPResults.ReadAll
  fFTPResults.Close
  
  oFTPScriptFSO.DeleteFile(sFTPTempFile)
  oFTPScriptFSO.DeleteFile (sFTPResults)
  
  If InStr(sResults, "226 Transfer complete.") > 0 Then
    FTPUpload = True
	msgbox FTPUpload
  ElseIf InStr(sResults, "File not found") > 0 Then
    FTPUpload = "Error: File Not Found"
	msgbox FTPUpload
  ElseIf InStr(sResults, "cannot log in.") > 0 Then
    FTPUpload = "Error: Login Failed."
	msgbox FTPUpload
  Else
    FTPUpload = "Error: Unknown."
  End If

  Set oFTPScriptFSO = Nothing
  Set oFTPScriptShell = Nothing
  App.MsgBox "File Uploaded Successfully","Confirmation",nlMsgInformation,nlMsgOKOnly
End Function


Sub QueryDatabase()
'Create a new instance of the RecordSet Object
	Set objRecordSet = CreateObject("ADODB.Recordset")
    strUserID = App.GetValue("varUserName")
	'Define your SQL statement
	objRecordSet.Source = "select usrmbr from hthdatv1.sysusrp where usrid='"&strUserID&"'"
	Set objRecordSet.ActiveConnection = objConnection
	objRecordSet.Open
	Do Until objRecordSet.EOF = True
	 mbr = objRecordSet.Fields("usrmbr").Value
	objRecordSet.MoveNext
	Loop
	objRecordSet.Close
	Set objRecordSet = Nothing 
End Sub


Function ConvertDate(strDate)
	Dim strYear
	Dim strMonth
	Dim strDay
	Dim strHour
	Dim strMinute
	Dim strSecond
	
	strYear = DatePart("yyyy", strDate)
	strMonth = DatePart("m", strDate)
	strDay = DatePart("d", strDate)
	strHour = DatePart("h", strDate)
	strMinute = DatePart("n", strDate)
	strSecond = DatePart("s", strDate)
	
	'If month is 1 digit (i.e. 4) then prefix it with a leading zero (i.e. 04)
	if Len(strMonth) = 1 Then
		strMonth = "0" & strMonth
	End If
	
	'If day is 1 digit (i.e. 4) then prefix it with a leading zero (i.e. 04)
	if Len(strDay) = 1 Then
		strDay = "0" & strDay
	End If

	'If hour is 1 digit (i.e. 4) then prefix it with a leading zero (i.e. 04)
	if Len(strHour) = 1 Then
		strHour = "0" & strHour
	End If

	'If minute is 1 digit (i.e. 4) then prefix it with a leading zero (i.e. 04)
	if Len(strMinute) = 1 Then
		strMinute = "0" & strMinute
	End If

	'If second is 1 digit (i.e. 4) then prefix it with a leading zero (i.e. 04)
	if Len(strSecond) = 1 Then
		strSecond = "0" & strSecond
	End If
	
	'Return the converted date (now in YYYYMMDDHNS format) to the calling function
	ConvertDate = strYear & strMonth & strDay & strHour & strMinute & strSecond
	
	
End Function
