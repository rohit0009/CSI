<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
    </head>
    <%
        if session("empid") = "" then 
                    response.redirect("sessionTimedOut.asp")
                    end if
        
        
        cust_id = request.querystring("cust_id")

       
            attribute = request.querystring("attribute")
            changedValue = request.querystring("changedValue")

            response.write cust_id & attribute & changedValue
            Dim Connection
	        Dim Recordset
            Dim RecordsetS
	        Dim SQL
            Set Connection = Server.CreateObject("ADODB.Connection")
	        Set Recordset = Server.CreateObject("ADODB.Recordset")
            
            Connection.Open "DSN=csidsn"
            
            SQL = ""
            SQL = "update customer set " & attribute & " = '" & changedValue & "' where cust_id="&cust_id&";"
	                    
            Recordset.Open SQL,Connection
            'Connection.Execute SQL,recaffected
            
            response.write "0"
            Set RecordsetS=nothing
            Set Recordset=nothing
            Set Connection=nothing

        
        
    %>
    <body>
        
    </body>
</html>
