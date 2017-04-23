<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Inventory</title>
        <style>
                #wrapper {
                position: absolute;
                width: 1400px;
                height: 750px;
                top: 0%;
                left: 0%;
                background-image: url('images/windows-central-background.jpg')
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
                border-radius: 15px;
                top: 220px;
                left: 30px;
                
            }
            #updatedelete
            {
                position: relative;
            }
            
            
            

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
           
            
             /* CSSTerm.com Simple Horizontal DropDown CSS menu */

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

            
        </style>
        <script language="JavaScript">
            function update()
            {
                document.F1.method = "post";
                document.F1.action = "selectItem.asp";
                document.F1.submit();
            }
            function insert()
            {
                document.F1.method = "post";
                document.F1.action = "insertItem2.asp";
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
                  <li><a class="active">Inventory</a></li>
                  <li><a href="customer0.asp">Customers</a></li>
                  <li><a href="supplier0.asp">Suppliers</a></li>
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
	                SQL = "SELECT * FROM inventory;"
	
	                Recordset.Open SQL,Connection                    
                    response.write "<table ><tr><th>Item No</th><th>Date of last order</th><th>Item Name</th><th>Stock Location</th><th>Describtion</th><th>Cost per Item</th><th>Stock Quantity</th><th>Total Value</th><th>Item Discontinued?</th></tr>"
                    
                    
                    Do While NOT Recordset.Eof
                        
                    if Recordset("item_discontinued") = "n" then
                        if Recordset("stock_quantity") = 0 then

		                    response.write "<tr bgcolor=red><td><center>"&Recordset("itemno")&"</center></td><td><center>"&Recordset("date_of_last_order")&"</td><td><center>"&Recordset("item_name")&"</td><td><center>"&Recordset("stock_location")&"</td><td><center>"&Recordset("describtion")&"</td><td><center>"&Recordset("cost_per_item")&"</td><td><center>"&Recordset("stock_quantity")&"</td><td><center>"&Recordset("total_value")&"</td>"
                            if Recordset("item_discontinued") = "y" then
                                response.write "<td><center>YES</td></tr>"
                            else
                                response.write "<td><center>NO</td></tr>"
                            end if
                        elseif Recordset("stock_quantity") < 6 then

		                    response.write "<tr bgcolor=#ff00dc><td><center>"&Recordset("itemno")&"</center></td><td><center>"&Recordset("date_of_last_order")&"</td><td><center>"&Recordset("item_name")&"</td><td><center>"&Recordset("stock_location")&"</td><td><center>"&Recordset("describtion")&"</td><td><center>"&Recordset("cost_per_item")&"</td><td><center>"&Recordset("stock_quantity")&"</td><td><center>"&Recordset("total_value")&"</td>"
                            if Recordset("item_discontinued") = "y" then
                                response.write "<td><center>YES</td></tr>"
                            else
                                response.write "<td><center>NO</td></tr>"
                            end if
                        else
                            response.write "<tr><td><center>"&Recordset("itemno")&"</center></td><td><center>"&Recordset("date_of_last_order")&"</td><td><center>"&Recordset("item_name")&"</td><td><center>"&Recordset("stock_location")&"</td><td><center>"&Recordset("describtion")&"</td><td><center>"&Recordset("cost_per_item")&"</td><td><center>"&Recordset("stock_quantity")&"</td><td><center>"&Recordset("total_value")&"</td>"
                            if Recordset("item_discontinued") = "y" then
                                response.write "<td><center>YES</td></tr>"
                            else
                                response.write "<td><center>NO</td></tr>"
                            end if
                        end if

	                    Recordset.MoveNext
                    else
                        Recordset.MoveNext
                    end if
                Loop
                
                
                %>
                
            </div>
            <div id="updatedelete">
                <form name="F1">
                    Operations
                    <input type="button" value="Insert" onclick="insert()">&nbsp&nbsp
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