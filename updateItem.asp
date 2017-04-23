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
        
        
        itemno = request.querystring("itemno")

       
            attribute = request.querystring("attribute")
            changedValue = request.querystring("changedValue")

            response.write itemno & attribute & changedValue
            Dim Connection
	        Dim Recordset
            Dim RecordsetS
	        Dim SQL
            Set Connection = Server.CreateObject("ADODB.Connection")
	        Set Recordset = Server.CreateObject("ADODB.Recordset")
            Set RecordsetS = Server.CreateObject("ADODB.Recordset")
            Connection.Open "DSN=csidsn"
            SQL = "select * from inventory where itemno = "&itemno
            RecordsetS.Open SQL,Connection
            SQL = ""
            SQL1 = ""
            if attribute = "cost_per_item" then 
	            SQL = "update inventory set " & attribute & " = " & changedValue & " where itemno="&itemno&";"
                newtotalvalue = changedValue * RecordsetS("stock_quantity")
                SQL1 = "update inventory set total_value = "& newtotalvalue & " where itemno = " &itemno& ";"
                Recordset.Open SQL1,Connection
            elseif attribute = "stock_quantity" then
                tempvalue = RecordsetS("stock_quantity") + changedValue
                SQL = "update inventory set " & attribute & " = " & tempvalue & " where itemno="&itemno&";"
                newtotalvalue = tempvalue * RecordsetS("cost_per_item")
                SQL1 = "update inventory set total_value = "& newtotalvalue & " where itemno = " &itemno& ";"
                Recordset.Open SQL1,Connection
            else
                SQL = "update inventory set " & attribute & " = '" & changedValue & "' where itemno="&itemno&";"
            end if
            Recordset.Open SQL,Connection
            
            response.write "0"
            Set RecordsetS=nothing
            Set Recordset=nothing
            Set Connection=nothing

        
        
    %>
    <body>
        
    </body>
</html>
