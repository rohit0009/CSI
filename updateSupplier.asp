
    <%

        
        if session("empid") = "" then 
                    response.redirect("sessionTimedOut.asp")
                    end if
        sid = request.querystring("sid")

       
            attribute = request.querystring("attribute")
            changedValue = request.querystring("changedValue")

            response.write sid & attribute & changedValue
            Dim Connection
	        Dim Recordset
            Dim RecordsetS
	        Dim SQL
            Set Connection = Server.CreateObject("ADODB.Connection")
	        Set Recordset = Server.CreateObject("ADODB.Recordset")
            
            Connection.Open "DSN=csidsn"
            
            SQL = ""
            SQL = "update supplier set " & attribute & " = '" & changedValue & "' where sid="&sid&";"
	                    
            Recordset.Open SQL,Connection
            'Connection.Execute SQL,recaffected
            
            response.write "0"
            Set RecordsetS=nothing
            Set Recordset=nothing
            Set Connection=nothing

        
        
    %>

