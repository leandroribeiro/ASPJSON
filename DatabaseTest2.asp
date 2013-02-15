<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #include file="include/clsADOHelper.asp" -->
<%
    Dim oADO, objRS
    
    Set oADO = new clsADOHelper
    Set objRS = oADO.GetRecordSetFromSQLString("SELECT TOP 10 codigo, nomeusu FROM usuario")

    do while not objRS.eof

        Response.Write(objRS("codigo"))
        Response.Write("<br/>")

        objRS.MoveNext
    loop

    oADO.CloseRecordSet()

%>
<html>
<head>
    <title></title>
</head>
<body>
    <h2>Database Test 2</h2>
</body>
</html>