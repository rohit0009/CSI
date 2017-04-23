<%
        
    if session("empid") = "" then 
    response.redirect("sessionTimedOut.asp")
    end if
    Dim Connection
	Dim Recordset
	Dim SQL

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")
    Connection.Open "DSN=csidsn"

    SQL = "insert into bill (cust_id,empid,date) values ("&session("cust_id")&","&session("empid")&",'"& request.querystring("date") &"');"
    Recordset.open SQL,Connection
    session.contents.remove("cust_id")
    date1 = request.querystring("date")

    set recordset = nothing
    Set Recordsetbill = Server.CreateObject("ADODB.Recordset")
	Set Recordsetlistitem = Server.CreateObject("ADODB.Recordset")
    Set Recordsetorders = Server.CreateObject("ADODB.Recordset")
    
 

    SQL = "select * from bill order by bid desc limit 1;"
    Recordsetbill.open SQL,Connection

    SQL1 = "select * FROM orders;"
    Recordsetorders.open SQL1,connection

    do while not Recordsetorders.eof

        Set Recordsetinventory = Server.CreateObject("ADODB.Recordset")
        SQL5 = "select * from inventory where itemno ="&Recordsetorders("itemno")&";"
        Recordsetinventory.open SQL5,Connection    

        Set Recordsetinventoryupdate = Server.CreateObject("ADODB.Recordset")
        changedValue = Recordsetinventory("stock_quantity") - Recordsetorders("quantity")
        changedValueTotal = Recordsetinventory("cost_per_item") * changedValue

        SQL6 = "update inventory set date_of_last_order= '"&date1 &"' where itemno="&Recordsetorders("itemno")
        Recordsetinventoryupdate.open SQL6,connection
        set Recordsetinventoryupdate = nothing
        set Recordsetinventory = nothing

        SQL3 = "insert into listitem (bid,itemno,quantity,total) values ("&Recordsetbill("bid")&","&Recordsetorders("itemno")&","&Recordsetorders("quantity")&","&Recordsetorders("total")&");"
        Recordsetlistitem.open SQL3,connection

        Recordsetorders.movenext
    loop


    set recordsetbill = nothing
    set recordsetlistitem = nothing
    set recordsetorders = nothing
    Set Recordsetorders = Server.CreateObject("ADODB.Recordset")
    SQL4 = "delete from orders;"
    Recordsetorders.open SQL4,connection
    set connection = nothing
    response.redirect("displayBill.asp")
%>
