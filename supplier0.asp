<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Customers</title>
        <style>
                #wrapper {
                position: absolute;
                width: 1400px;
                height: 750px;
                top: 0%;
                left: 0%;
                background-color: #98bcc1;
            }
            #head {
                overflow: hidden;
                position: relative;
                background-color: #0d1522;
                height: 90px;
                width: 1400px;
                top: 0%;
                left: 0%;
            }
            #title
			{
				position:absolute;
				top: 25%;
				left: 10%;
				color:#FFFFFF;
			}
            
            #inventory
            {
                position: absolute;   
                padding-left: 10px;
                padding-top: 20px;
                padding-right: 10px;
                padding-bottom: 10px;
                background-color: #81e4b7;
                height: max-content;
                top: 220px;
                left: 30px;
                border-radius: 15px;
                
            }
            #updatedelete
            {
                position: relative;
            }
            
            
            .drop_menu {
	            background:#005555;
	            padding:0;
	            margin:0;
	            list-style-type:none;
	            height:30px;
            }
            .drop_menu li { float:left; }
            .drop_menu li a {
	            padding:9px 20px;
	            display:block;
	            color:#fff;
	            text-decoration:none;
	            font:12px arial, verdana, sans-serif;
            }

            /* Submenu */
            .drop_menu ul {
	            position:absolute;
	            left:-9999px;
	            top:-9999px;
	            list-style-type:none;
            }
            .drop_menu li:hover { position:relative; background:#5FD367; }
            .drop_menu li:hover ul {
	            left:0px;
	            top:30px;
	            background:#5FD367;
	            padding:0px;
            }

            .drop_menu li:hover ul li a {
	            padding:5px;
	            display:block;
	            width:168px;
	            text-indent:15px;
	            background-color:#5FD367;
            }
.drop_menu li:hover ul li a:hover { background:#005555; }

            .active {
                background-color: #4CAF50;
                color: #111;
            }
            
            table, td, th {
                border: 2px solid #6b6a6a;
            }

            table {
                border-collapse: collapse;
                width: 100%;
            }

            th
            {
                height: 40px;
            }
            
            td {
                padding: 6px;
            }
            
        </style>
        
        <script language="JavaScript">
            function insert()
            {
                document.F2.method = "post";
                document.F2.action = "supplier1.asp";
                document.F2.submit();
            }
            function update()
            {
                document.F1.method = "post";
                document.F1.action = "selectSupplier.asp";
                document.F1.submit();
            }
        </script>
    </head>
    

<body>
        <div id="wrapper">
            
			<div id="head">
				<div id = "title" style="font-size:200%;">
					Computer Store Inventory Management System
				</div>
                
			</div>
            <div class="drop"> 
                <ul class="drop_menu">
                  <li><a href="Home.asp">Home</a></li>
                  <li><a href="Inventory.asp">Inventory</a></li>
                  <li><a href="customer0.asp">Customers</a></li>
                  <li><a class="active">Suppliers</a></li>
                  <li><a >Order</a>
                  <ul>
                        <li><a href="manageOrderAdd.asp">Add</a></li>
                        <li><a href="manageOrderDelete.asp">Delete</a></li>
                        <li><a href="manageOrders.asp">Orders</a></li>
                </ul>
                </li>
                  <li><a href="aboutUs.html">About Us</a></li>
                </ul>
            </div>
                
            <div id="inventory">
                
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
	                SQL = "SELECT * FROM supplier;"
	                
                    
	                Recordset.Open SQL,Connection                    
                    response.write "<table border = 1><tr><th>Supplier ID</th><th>First Name</th><th>Last Name</th><th>Email</th><th>Phone</th><th>Address</th><th>City</th><th>State</th><th>Pincode</th><th>Item Supplied</th></tr>"
                    
                    Do While NOT Recordset.Eof
                        Set Recordsetinventory = Server.CreateObject("ADODB.Recordset")
                        Recordsetinventory.open "select item_name from inventory where itemno = "&Recordset("itemno"),Connection
                      response.write "<tr><td><center>"&Recordset("sid")&"</center></td><td><center>"&Recordset("f_name")&"</td><td><center>"&Recordset("l_name")&"</td><td><center>"&Recordset("email")&"</td><td><center>"&Recordset("phone")&"</td><td><center>"&Recordset("address")&"</td><td><center>"&Recordset("city")&"</td><td><center>"&Recordset("state")&"</td><td><center>"&Recordset("pincode")&"</td><td>"&Recordsetinventory("item_name")&"</td></tr>"
                      set recordsetinventory =nothing
                      Recordset.MoveNext
	                Loop
                
                %>
                
            </div>
            <div id="updatedelete">
                Operations
            <form name="F2">
                <input type="button" value="Add New Supplier" onclick="insert()">&nbsp&nbsp
            </form>
            <form name="F1">
                
                <input type="button" value="Update" onclick="update()"></br></br>
            </form>
            </div>
        </div>
        
</body>
</html>
<%
    Recordset.Close
	Set Recordset=nothing
	Connection.Close
	Set Connection=nothing
%>