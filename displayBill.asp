<head>
    <style>
        #wrapper
        {
            position: absolute;
            top: 0px;
            left: 0px;
            background-color: CadetBlue;
            width: 976px;
            font-size: large;
            height: 476px;
            padding: 5px;
            
            border: 7px solid #000;
            
        }
        #title
        {
            position: relative;
            left: 30%;
            font-size: larger;
            top: 2%;
            width: 300px;
        }
        #custname
        {
            
            position: relative;
            left:55%;
        }
        #emp
        {
            position: relative;
            left: 70%;
        }
        #date
        {
            position: relative;
            left: 75%;
        }
        #listitem
        {
            position: relative;
            
            left: 5%;
        }
        table, td, th {
        border: 1px solid #6b6a6a;
        }

        table {
            border-collapse: collapse;
            width: 70%;
        }

        
        th,td {
            padding: 6px;
        }
    </style>
</head>    
<%
    if session("empid") = "" then 
    response.redirect("sessionTimedOut.asp")
    end if
    Dim Connection
	Dim Recordset
	Dim SQL

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set RecordsetBill = Server.CreateObject("ADODB.Recordset")
    
    Set Recordsetcustomer = Server.CreateObject("ADODB.Recordset")
    Set Recordsetemployee = Server.CreateObject("ADODB.Recordset")
	Connection.Open "DSN=csidsn"    
    
    
    SQL = "select * from bill order by bid desc limit 1"
    
    RecordsetBill.open SQL,Connection
    Recordsetemployee.open "select fname,lname from employee where empid="&RecordsetBill("empid"),connection
    do while not RecordsetBill.eof
        bid = RecordsetBill("bid")
        
        Recordsetcustomer.open "select fname,lname from customer where cust_id = "&RecordsetBill("cust_id"),connection
        content = content & "<div id = 'wrapper'><div id=title><center>PRN Computers Pvt Limited</center></div></br>-----------------------------------------------------------------------------------------------------------------------------------------------------------------</br>"
        content = content & "Bill Number : "&bid&"<span id=custname>Customer Name : "&Recordsetcustomer("fname")&" "&Recordsetcustomer("lname")&"</span><div id=date>Date : "&date&"</div><div id=emp>Employee Name : "&Recordsetemployee("fname")&" "&Recordsetemployee("lname")&"</div>"
        content = content & "</br></br><div id = 'listitem'><table ><caption style=font-size:25px;>Product Summary</caption><tr><th><center>Serial No</th><th><center>Item Name</th><th><center>Quantity</th><th><center>Rate</th><th><center>Sub Total</th></tr>"
        Set Recordsetlistitem = Server.CreateObject("ADODB.Recordset")
        Recordsetlistitem.open "select * from listitem where bid = "&bid,Connection
        count = 1
        total = 0
        do while not Recordsetlistitem.eof
            Set Recordsetinventory = Server.CreateObject("ADODB.Recordset")
            subtotal = Recordsetlistitem("total")
            Recordsetinventory.open "select * from inventory where itemno = "&Recordsetlistitem("itemno"),connection
            content = content & "<tr><td><center>"&count&"</td><td><center>"&Recordsetinventory("item_name")&"</td><td><center>"&Recordsetlistitem("quantity")&"</td><td><center>"&Recordsetinventory("cost_per_item")&"</td><td><center>"&subtotal&"</td></tr>"
            Recordsetlistitem.movenext
            total = total + subtotal
            count = count + 1
        loop
        set recordsetemployee = nothing
        Set Recordsetorder = nothing
        RecordsetBill.movenext
    loop
    content = content & "<tr></tr><tr></tr><tr></tr><tr></tr><tr></tr><tr><td colspan=4 align=right>Total</td><td><center>"&total&"/-</td></tr>"
    response.write content&"</table></div></br></br></br>-----------------------------------------------------------------------------------------------------------------------------------------------------------------</div>"


%>