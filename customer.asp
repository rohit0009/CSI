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
        fname = Request.Form("first_name")
        lname = Request.Form("last_name")
        email = Request.Form("email")
        phone = Request.Form("phone")
        address = Request.Form("address")
        city = Request.Form("city")
        state = Request.Form("state")
        pincode = Request.Form("zip")
        
        Dim Connection
	    Dim Recordset
	    Dim SQL

        Set Connection = Server.CreateObject("ADODB.Connection")
        Set Recordset = Server.CreateObject("ADODB.Recordset")
	    Connection.Open "DSN=csidsn"
                
        SQL = "select * from customer"
        Recordset.Open SQL,Connection
        
            Set Recordset = Server.CreateObject("ADODB.Recordset")
	        SQL = "insert into customer (fname,lname,email,phone,address,city,state,pincode) values ('"& fname &"','"&lname&"','"&email&"',"&phone&",'"&address&"','"&city&"','"&state&"','"&pincode&"');"
	
	        Connection.Execute SQL,recaffected
            if err<>0 then
                response.write "No permissions to write"
            else
                response.write "<head><title>Success</title></head><p style=font-size:50px;color:rgba(0, 255, 33, 0.90);>New Customer Added Successfully <a href=customer0.asp>Customer</a></p>"
                
            end if

    
	        Set Recordset=nothing
	
	    
        Set Connection=nothing

    %>

    <body>
        
    </body>
</html>
