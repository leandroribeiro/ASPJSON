<%
'##################################################################'
''
'Origem: http://code.google.com/p/aspjson/
'Versão Origem' 0.1.1
''
'Alterado Por: Leandro Ribeiro
'Alterado Em: 15/02/2013'
'Alteração: Adicionado método "RSToJSON"
''
'##################################################################'

Function QueryToJSON(dbc, sql)
        Dim rs, jsa
        Set rs = dbc.Execute(sql)
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null)(col.Name) = col.Value
                Next
        rs.MoveNext
        Wend
        Set QueryToJSON = jsa
End Function


Function RSToJSON(p_RS)
        Dim jsa
        Set jsa = jsArray()
        While Not (p_RS.EOF Or p_RS.BOF)
                Set jsa(Null) = jsObject()
                For Each col In p_RS.Fields
                        jsa(Null)(col.Name) = col.Value
                Next
        p_RS.MoveNext
        Wend
        Set RSToJSON = jsa
End Function

%>