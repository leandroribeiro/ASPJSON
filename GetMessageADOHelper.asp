<%
    'This will return the JSON object as a response
    'Comment below line to see this page result in browser
    Response.AddHeader "Content-Type", "application/json"
%>
<!-- #include file="include/clsADOHelper.asp" -->
<!-- #include file="include/JSON_2.0.4.asp" -->
<!-- #include file="include/JSON_UTIL_0.1.1.asp" -->
<%
    Dim oADO, objRS
    
    Set oADO = new clsADOHelper
    Set objRS = oADO.GetRecordSetFromSQLString("SELECT TOP 10 codigo, nomeusu FROM usuario")
    
    'This to make sure any un-related responses to be clear
    Response.Clear

    'This utility function in JSON_UTIL_0.1.1.asp accept SQL query and connection object and return JSON object. Extra Flush function directly write the response JSON object to response stream
    RSToJSON(objRS).Flush

    'Close database connnection
	oADO.CloseRecordSet()

%>