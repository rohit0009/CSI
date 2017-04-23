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
            #content
            {
                overflow:auto;
                position: absolute;
                
                
                top: 220px;
                left: 5%;
                background-color: #98bcc1;
               
            }
            #title
			{
				position:absolute;
				top: 25%;
				left: 10%;
				color:#FFFFFF;
			}
            
            #nav
            {
                width: inherit;
            }
            #nav ul {
                position: absolute;
                list-style-type: none;
                margin: 0;
                padding: 0;
                top: 90px;
                overflow: hidden;
                background-color: #2c42f3;
                width: inherit;
                
            }
            
            #nav li {
                float: left;
                
            }

            .active {
                background-color: #4CAF50;
                color: #111;
            }
            
            #nav li a:hover:not(.active) {
               background-color: #111;
               
            }
            #nav li a 
            {
                display: block;
                color: white;
                text-align: center;
                padding: 15px 30px;
                text-decoration: none;
                
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

       <script>
           function add() {
               alert("hi");
               document.getElementById("content").width = "500px";
               document.getElementById("content").innerHTML = '<object type = "text/html" data = selectCustomer.asp></object>';
           }   
       </script>
    </head>
    
<%
   if session("empid") = "" then 
    response.redirect("sessionTimedOut.asp")
    end if

    Qstr = Request.querystring("login")
    if Qstr = "logout" then
        
        dim con
        dim rs
        Set con = Server.CreateObject("ADODB.Connection")
	    Set rs = Server.CreateObject("ADODB.Recordset")
        con.Open "DSN=csidsn"
        SQL = "delete from orders where emp_id="&session("empid")
        rs.open SQL,con
        session.abandon
        set rs = nothing
        set con = nothing
        response.redirect("index.html")
    end if


    Dim Connection
	Dim Recordset
	Dim SQL

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")

    dim fname
    dim lname

	Connection.Open "DSN=csidsn"
	SQL = "SELECT fname,lname from employee where empid = "& session("empid") &";"
	
	Recordset.Open SQL,Connection
    fname = Recordset("fname")
    lname = Recordset("lname")

%>
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
                  <li><a href="supplier0.asp">Suppliers</a></li>
                  <li><a class="active">Order</a>
                  <ul>
                        <li><a href="manageOrderAdd.asp">Add</a></li>
                        <li><a href="manageOrderDelete.asp">Delete</a></li>
                        <li><a href="manageOrders.asp">Orders</a></li>
                </ul>
                </li>
                  <li><a href="aboutUs.html">About Us</a></li>
                </ul>
            </div>
        </div>
        
        <div id ="content">
           
            <iframe src="showCart.asp" width ="1000" height="500" scrolling ="no">
        </div>
        
</body>
</html>
<%
    Recordset.Close
	Set Recordset=nothing
	Connection.Close
	Set Connection=nothing
%>