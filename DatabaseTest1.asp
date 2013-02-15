<%@  language="VBScript" %>
<% Option Explicit %>
<%
    Const CONNECTION_STRING = "DRIVER={SQL Server};Server=LD-LEANDRO\SQLEXPRESS;Database=usuarios_apol;UID=sa;PWD=123"

    Dim objConn, objRS
    Dim strSQL

    Set objConn = Server.CreateObject("ADODB.Connection")
    
    objConn.ConnectionString = CONNECTION_STRING        
    
    If objConn.State <> 1 Then
        objConn.Open
    End If    

    'This query simply return first name and last name
    strSQL = "SELECT TOP 10 codigo, nomeusu FROM usuario"

    Set objRS = Server.CreateObject("ADODB.Recordset")
    objRS.Open strSQL, objConn

    do while not objRS.eof

        Response.Write(objRS("codigo"))
        Response.Write("<br/>")

        objRS.MoveNext
    loop
        
    If objConn.State <> 0 Then
            objConn.Close   
    End If

    Set objConn = Nothing
%>
<html>
<head>
    <title></title>
</head>
<body>
    <h2>Database Test Simple</h2>
</body>
</html>