<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8">
		<meta http-equiv="X-UA-Compatible" content="IE=edge">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<!-- The above 3 meta tags *must* come first in the head; any other head content must come *after* these tags -->
		<meta name="description" content="">
		<meta name="author" content="">

		<!-- Bootstrap core CSS -->
		<link href="bootstrap-3.3.7/docs/dist/css/bootstrap.min.css" rel="stylesheet">
		<!-- Bootstrap theme -->
		<link href="bootstrap-3.3.7/docs/dist/css/bootstrap-theme.min.css" rel="stylesheet">
		<!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
		<link href="bootstrap-3.3.7/docs/assets/css/ie10-viewport-bug-workaround.css" rel="stylesheet">

		

		<!-- Just for debugging purposes. Don't actually copy these 2 lines! -->
		<!--[if lt IE 9]><script src="../../assets/js/ie8-responsive-file-warning.js"></script><![endif]-->
		<script src="bootstrap-3.3.7/docs/assets/js/ie-emulation-modes-warning.js"></script>

		<!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
		<!--[if lt IE 9]>
		  <script src="https://oss.maxcdn.com/html5shiv/3.7.3/html5shiv.min.js"></script>
		  <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
		<![endif]-->
        <style>
            #wrapper {
                position: absolute;
                width: 1350px;
                height: 750px;
                top: 0%;
                left: 0%;
                background-image: url('images/bg1.jpg');
                
            }
            #head {
                overflow: hidden;
                position: relative;
                background-color: #0d1522;
                height: 90px;
                width: 1350px;
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
            #user
            {
                position:absolute;
				top:5%;
				right: 10%;
               
            }
            a.logout:visited 
            {
                color: #0000EE;
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

            
            .active {
                background-color: #4CAF50;
                color: #111;
            }
            
            
        </style>
       <script>
           
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
	
	Set Recordsetinventory = Server.CreateObject("ADODB.Recordset")
	Set Recordsetorder = Server.CreateObject("ADODB.Recordset")
	Recordsetorder.Open "select count(oid) as o from orders;",Connection
	Recordsetinventory.Open "select count(itemno) as c from inventory;",Connection

%>
<body>
        <div id="wrapper">
			<div id="head">
				<div id = "title" style="font-size:200%;">
					Computer Store Inventory Management System
				</div>
                
                <span id="user" style="color:#FFFFFF;"><h4> Welcome <% response.write fname &" "& lname%> <br/><br/> Do you want to <a class="logout" href="Home.asp?login=logout" style="font-size: 100%" >Logout?</a></h4></span>
			</div>
            <div class="drop"> 
                <ul class="drop_menu">
                  <li><a class="active">Home</a></li>
                  <li><a href="Inventory.asp">Inventory <span class="badge" style="background-color:black;"><%response.write Recordsetinventory("c")%></span></a></li>
                  <li><a href="customer0.asp">Customers</a></li>
                  <li><a href="supplier0.asp">Suppliers</a></li>
                  <li><a >Order</a>
                  <ul>
                        <li><a href="manageOrderAdd.asp">Add</a></li>
                        <li><a href="manageOrderDelete.asp">Delete</a></li>
                        <li><a href="manageOrders.asp">Orders <span class="badge" style="background-color:black;"><%response.write Recordsetorder("o")%></span></a></li>
                </ul>
                </li>
                  <li><a href="aboutUs.html">About Us</a></li>
                </ul>
            </div>
			<div class="container theme-showcase" role="main">
				<div class="page-header">
					<h1><strong>Welcome</strong></h1>
				</div>
				<div class="alert alert-success" role="alert">
					You successfully logged in.
				</div>
			</div>
        </div>
        
</body>
</html>
<%
	Recordsetinventory.Close
	Set Recordsetinventory=nothing
	Recordsetorder.Close
	Set Recordsetorder=nothing
    Recordset.Close
	Set Recordset=nothing
	Connection.Close
	Set Connection=nothing
%>