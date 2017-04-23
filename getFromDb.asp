
    <%
        if session("empid") = "" then 
        response.redirect("sessionTimedOut.asp")
        end if
        itemno = Request.querystring("itemno")
        Dim Connection
	    Dim Recordset
	    Dim SQL

	    Set Connection = Server.CreateObject("ADODB.Connection")
	    Set Recordset = Server.CreateObject("ADODB.Recordset")

	    Connection.Open "DSN=csidsn"
        Replace custname,"%20"," "
	    SQL = "SELECT * FROM inventory where itemno = " & itemno
	
	    Recordset.Open SQL,Connection
        Do While NOT Recordset.Eof
            response.write Recordset("cost_per_item") & " " & Recordset("stock_quantity")
            Recordset.movenext
        loop
        Set Recordset=nothing
	    Set Connection=nothing
    %>

