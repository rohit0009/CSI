<%
        if session("empid") = "" then 
        response.redirect("sessionTimedOut.asp")
        end if
    Dim Connection
	Dim Recordset
	Dim SQL

	Set Connection = Server.CreateObject("ADODB.Connection")
    Set Recordsetinventoryupdate = Server.CreateObject("ADODB.Recordset")
    
    Set Recordsetorder = Server.CreateObject("ADODB.Recordset")
	Connection.Open "DSN=csidsn"
    
    
    Recordsetorder.open "select * from orders where emp_id = "&session("empid"),connection
    do while not Recordsetorder.eof
        Set Recordsetinventory = Server.CreateObject("ADODB.Recordset")
        Recordsetinventory.open "select * from inventory where itemno = "&Recordsetorder("itemno"),connection
        changedValue = Recordsetinventory("stock_quantity") + Recordsetorder("quantity")
        changedValueTotal = Recordsetinventory("cost_per_item") * changedValue
        Recordsetinventoryupdate.open "update inventory set stock_quantity = "&changedValue&" , total_value = "&changedValueTotal&" where itemno = "&Recordsetorder("itemno"),connection


	    set recordsetinventory = nothing
        Recordsetorder.movenext
    loop

    Set Recordsetdelete = Server.CreateObject("ADODB.Recordset")
    SQL ="delete from orders where emp_id ="& session("empid")
    Recordsetdelete.open SQL,connection
    
    set connection = nothing
    set recordsetdelete = nothing
    
%>