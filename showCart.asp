
<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Homepage</title>
        <style>
            #wrapper {
                position: absolute;
                width: 1400px;
                height: 750px;
                top: 0%;
                left: 0%;
                background-color: rgba(128, 128, 128, 0.77);
            }
            
            #content
            {
                position: absolute;
                width: 800px;
                
                background-color: #ffd800;
                top: 10%;
                left:5%;
            }
            
            
        </style>
       <script>
           function genrateBill() {
               var dateObj = new Date();
               var month = dateObj.getUTCMonth() + 1; //months from 1-12
               var day = dateObj.getUTCDate();
               var year = dateObj.getUTCFullYear();
               var date = year + "-" + month + "-" + day;
               
               
               document.F3.method = "POST";
               document.F3.action = "processBill.asp?date=" + date;
               document.F3.submit();
               /*var xhttp = new XMLHttpRequest();
               var op = "";
               xhttp.onreadystatechange = function () {
               if (this.readyState == 4 && this.status == 200) {
               document.F3.metod = "Post";
               document.F3.action = "displayBill.asp";
               document.F3.submit();

               }
               };
               xhttp.open("POST", "processBill.asp?date="+date , false);
               xhttp.send();*/
           }
           function redirect() {
               document.F3.method = "POST";
               document.F3.action = "order.asp";
               document.F3.submit();
           }
           function cancel() {
               document.F3.method = "POST";
               document.F3.action = "deleteOrder.asp";
               document.F3.submit();
           }
       </script>
    </head>
    
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
	

%>
<body>
        <div id="wrapper">
            
            <%
                
                SQL = "select count(oid) as c from orders"
                Recordset.Open SQL,Connection
                
                
                
                count = 0
                if Recordset("c") = "0" then
                    content = content & "<div id = content> <center><h2 style='color:red;'> No Items Ordered </h2></br> <a href=selectCustomer.asp>Order Items First</a></center></div>"
                else
                    Set RecordsetCart = Server.CreateObject("ADODB.Recordset")
                    SQL1 = "select * from orders"
                    RecordsetCart.open SQL1,Connection

                    content = content & "<form name=F3>"
                    content = content & "<div id = content ><center><H2>Orders</h2><table border=2><tr><th>Serial No</th><th>Item Name</th><th>Quantity</th><th>Sub Total</th></tr>"
                    do while not RecordsetCart.eof
                        count = count + 1
                        Set RecordsetInventory = Server.CreateObject("ADODB.Recordset")
                        SQL2 = "select * from inventory where itemno = "&RecordsetCart("itemno")&";"
                        RecordsetInventory.open SQL2,Connection
                        content = content & "<tr><td>"&count&"</td><td>"&RecordsetInventory("item_name")&"</td><td>"&RecordsetCart("quantity")&"</td><td>"&RecordsetCart("total")&"</td></tr>"
                        RecordsetInventory.movenext
                        set RecordsetInventory = nothing
                        RecordsetCart.movenext
                    loop
                    content = content & "</table></br><input type = button style = 'color:red;'value='Order More' onclick = redirect()>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<input type= button value = 'Cancel Order' onclick = 'cancel()'>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<input style = 'background-color:#00ff21;' type=button value='Generate Bill' onclick=genrateBill()></center>"
                end if
                content = content & "</form></div>"
                response.write content
            %>

</div>        
        
</body>
</html>
<%
    
	Set Recordset=nothing
	Set RecordsetCart = nothing
	Set Connection=nothing
%>
