Option Explicit
' This solution uses the Microsoft ActiveX Data Objects x.x Library


dim objConnection
dim strTable 

Sub ConnectToDatabase()
    dim strConnectionString
    dim strServerName
    dim strDatabase 
    dim strUserID 
    dim strPassword

	strServerName = "Printers.Hi-Techhealth.com"
	strDatabase = "dbname"
	strUserID = "user"
	strPassword = "password"
    

    'Create a new instance of the Connection Object
    Set objConnection = CreateObject("ADODB.Connection")
	

	strConnectionString = "Provider=IBMDA400;Data Source=" & strServerName & ";User Id=" & strUserID & ";Password=" & strPassword
		strTable  = "SECNAM"
	'Connect to the database
	objConnection.Open strConnectionString
	App.MsgBox "Connection Successful", "Confirmation",nlMsgExclamation,nlMsgOKOnly 
End Sub

Sub QueryDatabase()
	ConnectToDatabase
	Dim objRecordSet
	Dim intListViewIndex
	Dim strSearch
	'The percent sign (i.e. %) is used to indicate a partial match.  Therefore, if the user leaves the states combobox blank all states will be returned.
	strSearch = Trim(App.ActiveForm.txtSearch.Text) & "%"
	strSearch = "Billing"
	'Create a new instance of the RecordSet Object
    Set objRecordSet = CreateObject("ADODB.Recordset")

	'Define your SQL statement
	objRecordSet.Source = "SELECT * " & _
			"FROM HTHDATV1/" & strTable  & " " & _
                "WHERE SPRGNM LIKE '%" & strSearch & "%' " 
	

	'Let your RecordSet know what connection it will be using
	Set objRecordSet.ActiveConnection = objConnection
    'Open your RecordSet (i.e. Execute the query)
    objRecordSet.Open
	
	App.OpenForm "Resultsearch"

	'Clear the ListView
    App.ActiveForm.lvwRS.ListItems.Clear
        
    intListViewIndex = 0
    App.ActiveForm.Message = "Record Count: " & intListViewIndex
	
	
	Do Until objRecordSet.EOF = True
		'Add the current record to the ListView
          App.ActiveForm.lvwRS.ListItems.Add intListViewIndex, objRecordSet.Fields("ID").Value & " "
          App.ActiveForm.lvwRS.ListItems(intListViewIndex).SubItems(0).Text = objRecordSet.Fields("SPRGID").Value
          App.ActiveForm.lvwRS.ListItems(intListViewIndex).SubItems(1).Text = objRecordSet.Fields("SMNUID").Value
          App.ActiveForm.lvwRS.ListItems(intListViewIndex).SubItems(2).Text = objRecordSet.Fields("SMNUNO").Value
          App.ActiveForm.lvwRS.ListItems(intListViewIndex).SubItems(3).Text = objRecordSet.Fields("SPRGNM").Value
		
		'Move to the next record in the RecordSet
        objRecordSet.MoveNext
        
        'Increment intListViewIndex since we will now be adding to the next row in the ListView.
        intListViewIndex = intListViewIndex + 1
        
        'We test to see if the last digit in intListViewIndex is 0 because we only want to display the record count in increments of 10 (i.e 10, 20, 30, etc).
        if right(intListViewIndex, 1) = "0" then
        	App.ActiveForm.Message = "Record Count: " & intListViewIndex
        end if
		
	Loop
	
	App.ActiveForm.Message = "Record Count: " & intListViewIndex
	'Close the RecordSet
    objRecordSet.Close
    'Destroy the objRecordSet object variable from memory
    Set objRecordSet = Nothing
	DisConnectFromDatabase
End Sub
	

Sub DisConnectFromDatabase()
'Disconnect from the database

    'Close the Database Connection
    objConnection.Close
    'Destroy the objConnection object variable from memory
    Set objConnection = Nothing
    
End Sub
