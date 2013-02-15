<%
    'This will return the JSON object as a response
    'Comment below line to see this page result in browser
    Response.AddHeader "Content-Type", "application/json"
%>
<!-- #include file="include/JSON_2.0.4.asp" -->
<!-- #include file="include/JSON_UTIL_0.1.1.asp" -->
<%
    Const CONNECTION_STRING = "DRIVER={SQL Server};Server=LD-LEANDRO\SQLEXPRESS;Database=usuarios_apol;UID=sa;PWD=123"

    Dim objConn
    Dim strSQL

    Set objConn = Server.CreateObject("ADODB.Connection")
    
    objConn.ConnectionString = CONNECTION_STRING        
    
    If objConn.State <> 1 Then
        objConn.Open
    End If    

    'This is just a helper function to open database connection
    'Call fnOpenDBConnection("dbName",Com)

    'This query simply return first name and last name
    strSQL = "SELECT TOP 10 codigo, nomeusu FROM usuario"
    
    'This to make sure any un-related responses to be clear
    Response.Clear

    'This utility function in JSON_UTIL_0.1.1.asp accept SQL query and connection object and return JSON object. Extra Flush function directly write the response JSON object to response stream
    QueryToJSON(objConn, strSQL).Flush

    'Close database connnection
    'Call fnCloseDBConnection(Com)

    If objConn.State <> 0 Then
            objConn.Close   
    End If
    Set objConn = Nothing

%>