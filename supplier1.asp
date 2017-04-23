<!DOCTYPE html>

<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>

        <script language="JavaScript">
            function validate() {
                var fname = document.forms["F1"]["first_name"].value;
                var lname = document.forms["F1"]["last_name"].value;
                var email = document.forms["F1"]["email"].value;
                var phone = document.forms["F1"]["phone"].value;
                var city = document.forms["F1"]["city"].value;
                var state = document.forms["F1"]["state"].value;
                var zip = document.forms["F1"]["zip"].value;
                var iname = document.forms["F1"]["iname"].value;

                
                var err = "";
                var f = 0;

                if (fname == "" || lname == "" || email == "" || phone == "" || city == "" || state == "" || zip == "" || iname == "") {
                    f = 1;
                    err = err + "Please fill all the above fields";
                }

                else {
                    if ((/[0-9]+/.test(fname))) {
                        f = 1;
                        err = err + "Please enter a valid format. "
                    }

                    if ((/[0-9]+/.test(lname))) {
                        f = 1;
                        err = err + "Please enter a valid format. "
                    }

                    if (!(/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(email))) {
                        f = 1;
                        err = err + " Email id is not valid.";
                    }

                    if ((/[0-9]+/.test(city))) {
                        f = 1;
                        err = err + "Please enter a valid format. "
                    }

                    if ((/[0-9]+/.test(state))) {
                        f = 1;
                        err = err + "Please enter a valid format. "
                    }

                    if (!(/[0-9]+/.test(phone))) {
                        f = 1;
                        err = err + "Please enter a valid number. "
                    }

                    if (!(/[0-9]+/.test(zip))) {
                        f = 1;
                        err = err + "Please enter a valid pincode."
                    }


                }

                if (f == 1) {
                    alert(err);
                    return false;
                }
                return true;
            }
        </script>

    </head>

        
       
<head>
 <style>
                #wrapper {
                position: absolute;
                width: 1400px;
                height: 750px;
                top: 0%;
                left: 0%;
                background-image: url('images/main-bg.jpg');
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
            
            
            #updatedelete
            {
                position: relative;
            }
            
            

            .active {
                background-color: #4CAF50;
                color: #111;
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
        </style>
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





<div class="container">

    <form name ="F1" class="well form-horizontal" action="supplier.asp" method="post"  id="contact_form">


<!-- Form Name -->


<!-- Text input-->


<div class="form-group">
    <br/><br/>
    Enter Supplier Details
    <br/><br/>
  <label class="col-md-4 control-label">First Name</label>
  <div class="col-md-4 inputGroupContainer">
  <div class="input-group">
  <span class="input-group-addon"><i class="glyphicon glyphicon-user"></i></span>
  <input  name="first_name" placeholder="Enter First Name" class="form-control"  type="text">
    </div>
  </div>
</div>

<!-- Text input-->

<div class="form-group">
  <label class="col-md-4 control-label" >Last Name</label> 
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
  <span class="input-group-addon"><i class="glyphicon glyphicon-user"></i></span>
  <input name="last_name" placeholder="Enter Last Name" class="form-control"  type="text">
    </div>
  </div>
</div>

<!-- Text input-->
       <div class="form-group">
  <label class="col-md-4 control-label">E-Mail</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-envelope"></i></span>
  <input name="email" placeholder="Enter E-Mail Address" class="form-control"  type="text">
    </div>
  </div>
</div>


<!-- Text input-->
       
<div class="form-group">
  <label class="col-md-4 control-label">Phone #</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-earphone"></i></span>
  <input name="phone" placeholder="Enter phone no" class="form-control" type="text">
    </div>
  </div>
</div>

<!-- Text input-->
      
<div class="form-group">
  <label class="col-md-4 control-label">Address</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-home"></i></span>
  <input name="address" placeholder="Enter Address" class="form-control" type="text">
    </div>
  </div>
</div>

<!-- Text input-->
 
<div class="form-group">
  <label class="col-md-4 control-label">City</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-home"></i></span>
  <input name="city" placeholder="Enter city" class="form-control"  type="text">
    </div>
  </div>
</div>

<!-- Select Basic -->
   
<div class="form-group"> 
  <label class="col-md-4 control-label">State</label>
    <div class="col-md-4 selectContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-list"></i></span>
    <select name="state" class="form-control selectpicker" >
      <option value=" " >Please select your state</option>
      <option>Andhra Pradesh </option>
      <option>Arunachal Pradesh</option>
      <option >Assam</option>
      <option >Bihar</option>
      <option >Chhattisgarh</option>
      <option >Goa</option>
      <option >Gujrat</option>
      <option >Haryana</option>
      <option >Himachal Pradesh</option>
      <option> Jammu & Kashmir</option>
      <option >Jharkhand</option>
      <option >Karnataka</option>
      <option >Kerala</option>
      <option >Madhya Pradesh</option>
      <option >Maharashtra</option>
      <option >Manipur</option>
      <option> Meghalaya</option>
      <option >Mizoram</option>
      <option >Nagaland</option>
      <option>Odisha</option>
      <option >Punjab</option>
      <option> Sikkim</option>
      <option >Tamil Nadu</option>
      <option >Telangana</option>
      <option>Tripura</option>
      <option>Uttarakhand</option>
      <option>Uttar Pradesh</option>
      <option>West Bengal</option>
      
    </select>
  </div>
</div>
</div>

<!-- Text input-->

<div class="form-group">
  <label class="col-md-4 control-label">Zip Code</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-home"></i></span>
  <input name="zip" placeholder="Enter Pin Code" class="form-control"  type="text">
    </div>
</div>
</div>
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
	SQL = "select * from inventory where item_discontinued='n'"
	
	Recordset.Open SQL,Connection
%>
<div class="form-group">
  <label class="col-md-4 control-label">Item Name</label>  
    <div class="col-md-4 inputGroupContainer">
    <div class="input-group">
        <span class="input-group-addon"><i class="glyphicon glyphicon-home"></i></span>
  <select name="iname" class="form-control">
  <option value=" ">Please select your option</option>
    <%Do While NOT Recordset.Eof
        response.write "<option value="&Recordset("itemno")&">"&Recordset("item_name")&"</option>"
        Recordset.movenext
    loop
    set connection = nothing
    set recordset = nothing
    %>
  </select>
    </div>
</div>
</div>

<!-- Text area -->
 

<!-- Success message -->
<div class="alert alert-success" role="alert" id="success_message">Success <i class="glyphicon glyphicon-thumbs-up"></i> Thanks for contacting us, we will get back to you shortly.</div>

<!-- Button -->
<div class="form-group">
  <label class="col-md-4 control-label"></label>
  <div class="col-md-4">
    <button type="submit" onclick="return validate()" class="btn btn-warning" >Confirm <span class="glyphicon glyphicon-send"></span></button>
  <input type="reset" value="Clear">  

</div>
</div>


</form>
</div>
    </div><!-- /.container -->

    </body>
</html>
