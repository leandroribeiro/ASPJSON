<%

'References: http://www.sitepoint.com/forums/showthread.php?491770-Build-a-Database-Connections-Class-in-Classic-ASP

Const CONNECTION_STRING = "DRIVER={SQL Server};Server=LD-LEANDRO\SQLEXPRESS;Database=usuarios_apol;UID=sa;PWD=123"

Class clsADOHelper

	'#############  Attributes ##############

	Private strConnection	'## Connection string (change depending on what system we are using)
	Private objConn			'## Connection object
	Private objComm			'## Command Object
	Private objRS			'## Recordset object
	
	'#############  Properties ##############

	Public Property Get ObjectConn()
			ObjectConn = objConn
		End Property	

	'Public Property Let SwitchConnection(ByRef strConn)
		'## Will allow user to change the connection from the default set up	
		'strConnection = strConn
		'Call SwitchConnection(strConnection)
	'End Property

	'#############  Constructors and Destructors ##############

	Private Sub Class_Initialize()
		'## What happens when the class is opened
		'strConnection = ""
		strConnection = CONNECTION_STRING

		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.ConnectionString = strConnection
	End Sub
	
	Private Sub Class_Terminate()
		'## What happens when the class is closed
		
		'## Close connections
		If objConn.State <> 0 Then
			objConn.Close	
		End If
		Set objConn = Nothing
	End Sub			

	'#############  Public Sub and Functions, accessible to the web pages ##############

	Public Sub SQLExecuteFromSQLString(ByRef strSQL)
		'## Execute code and return nothing
		If objConn.State <> 0 Then
			objConn.Close
		End If	
		
		objConn.Execute strSQL		
	End Sub

	'## This replicates the .NET ExecuteScalar
	Public Function ExecuteScalarFromSQLString(ByRef sSQL)
		'## This is used when passing back single results. Replicating a .NET piece of functionality
		Dim objScalar
		Set objScalar = GetRecordSetFromSQLString(sSQL)
		
		If Not objScalar.EOF Then
			ExecuteScalar = objScalar(0)
		Else
			'## Nothing returned
			ExecuteScalar = -1
		End If
		
		CloseRecordSet()
	End Function 'ExecuteScalar
	
	Public Function GetRecordSetFromSQLString(ByRef strRS)
		If objConn.State <> 1 Then
			objConn.Open
		End If
	
		Set objRS = Server.CreateObject("ADODB.Recordset")
	
		objRS.Open strRS, objConn
		
		Set GetRecordSetFromSQLString = objRS
	End Function
	
	'## Using SP code within class
	'##########################################################################
	Public Sub CallSPNeedParams(ByRef strStoredProc)
		If objConn.State <> 1 Then
			objConn.Open
		End If
		
		If Not IsObject(objComm) Then
			Set objComm = Server.CreateObject("ADODB.Command") '## This will be used for Stored Procedures
		End If
		
		With objComm
			.ActiveConnection = objConn
			.CommandText = strStoredProc
			.CommandType = adCmdStoredProc
		End With
	
		If Not IsObject(objRS) Then
			Set objRS = Server.CreateObject("ADODB.Recordset")
		End If
			
		Set objRS.ActiveConnection = objConn '## Set connection
		Set objRS.Source = objComm'' Set source to use command object			
	End Sub
	
	Public Sub ApendParamsToRecordSet(ByRef Name, ByRef TypeParam, ByRef Direction, ByRef Size, ByRef Value)
		'Type adDate adDBDate, adVarChar, adChar, adBoolean
		If IsObject(objComm) Then
			objComm.Parameters.Append objComm.CreateParameter(Name, TypeParam, Direction, Size, Value)			
		End If
	End Sub
	
	Public Function GetRecordSetSPParams(ByRef strStoredProc)
		If strStoredProc = objComm.CommandText Then
			'## This is being called for the right SP
			objRS.Open
			Set GetRecordSetSPParams = objRS
			
			'## Need to clear out params from Command object
			Do While (objComm.Parameters.Count > 0)
				objComm.Parameters.Delete 0
			Loop
			
		End If
	End Function
	
	Public Function ExecuteScalarSetSPParams(ByRef strStoredProc)
		'## This is used when passing back single results. Replicating a .NET piece of functionality
		If strStoredProc = objComm.CommandText Then		
			objRS.Open
			If Not objRS.EOF Then
				ExecuteScalar = objRS(0)
			Else
				'## Nothing returned
				ExecuteScalar = -1
			End If
		
			CloseRecordSet()
		End If
	End Function 'ExecuteScalar				
	
	Public Sub ExecuteSPButNoRecordsReturned(ByRef strStoredProc)
		If strStoredProc = objComm.CommandText Then
			objComm.Execute
		End If
	End Sub	'ExecuteSPButNoRecordsReturned()
	
	'#############  Connections Subs, accessible to the web pages ##############

	Public Sub CloseRecordSet()	
		If objRS.State <> 0 Then
			objRS.Close	
		End If
		
		Set objRS = Nothing
	End Sub

	Public Sub CloseCommObject()
		If IsObject(objComm) Then
			Set objComm = Nothing
		End If
	End Sub		
	
	
	Public Sub SwitchConnection(ByRef strConn)
		'## Will allow user to change the connection from the default set up	
		strConnection = strConn
		If objConn.State <> adStateClosed Then
			objConn.ConnectionString = strConnection
		End If
	End Sub

End Class 'clsADOHelper
%>