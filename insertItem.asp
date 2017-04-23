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
        dim day
        dim month
        dim year

        iname = Request.form("iname")
        
        stockL = Request.form("stockL")
        desc = Request.form("desc")
        cost = Request.form("cost")
        qty = Request.form("qty")

        day = Request.QueryString("day")
        month = Request.QueryString("month")
        year = Request.QueryString("year")

        

        Dim Connection
    	Dim Recordset
	    Dim SQL

	    Set Connection = Server.CreateObject("ADODB.Connection")
        Set Recordset = Server.CreateObject("ADODB.Recordset")
	    Connection.Open "DSN=csidsn"

        total_value = cost * qty
        SQL = "select * from inventory"
        Recordset.Open SQL,Connection
        
            Set Recordset = Server.CreateObject("ADODB.Recordset")
	        SQL = "insert into inventory (item_name,stock_location,describtion,cost_per_item,stock_quantity,total_value,item_discontinued) values ('"& iname &"','"&stockL&"','"&desc&"',"&cost&","&qty&","&total_value&",'n');"
	        
	        Connection.Execute SQL,recaffected
            if err<>0 then
                response.write "No permissions to write"
            else
                response.write "<head><title>Success</title></head><p style=font-size:50px;color:rgba(0, 255, 33, 0.90);>New item inserted successfully! <a href=inventory.asp>Take me back</a></p>"
                
            end if

	    Set Recordset=nothing 
        Set Connection=nothing
    %>

    <body>
        
    </body>
</html>
