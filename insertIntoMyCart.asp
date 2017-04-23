<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
    </head>
    <%
        if session("empid") = "" then
            response.write "1"
        else
            itemno = request.querystring("itemno")
            ordered_quantity = request.querystring("ordered_quantity")
            cost_per_item = request.querystring("cost_per_item")
            Dim Connection
	        Dim Recordset
	        Dim SQL

	        Set Connection = Server.CreateObject("ADODB.Connection")
	                
            Set Recordset = Server.CreateObject("ADODB.Recordset")

	        Connection.Open "DSN=csidsn"
	                
            SQL = "insert into orders (itemno,quantity,cust_id,emp_id,cost_per_item,total) values ("&itemno&","&ordered_quantity&","&session("cust_id")&","&session("empid")&","&cost_per_item&","&ordered_quantity * cost_per_item&");"
            
            Recordset.Open SQL,Connection

            set Recordset = nothing
            Set RecordsetInventory = Server.CreateObject("ADODB.Recordset")
            
            RecordsetInventory.open "select * from inventory where itemno = "&itemno,Connection

            changedValue = RecordsetInventory("stock_quantity") - ordered_quantity
            changedValueTotal = RecordsetInventory("cost_per_item") * changedValue
            Set RecordsetInventoryUpdate = Server.CreateObject("ADODB.Recordset")
            RecordsetInventoryUpdate.open "update inventory set stock_quantity = "&changedValue&", total_value = "&changedValueTotal&" where itemno = "&itemno,Connection

            set recordsetInventory = nothing
            set RecordsetInventoryUpdate = nothing
            set Connection = nothing
        end if
            
    %>
    <body>
        
    </body>
</html>
