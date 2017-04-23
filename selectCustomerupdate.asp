<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
        <style>
            body
            {
                background: url("images/main-bg.jpg");
            }
            #content
            {
                position: absolute;
                top: 10%;
                left:5%;
                background-color: rgba(235, 165, 165, 0.48);
                padding: 20px;
            }
            #links
            {
                position: absolute;
                top: 10%;
                right: 5%;
                
            }
            #displaytable table,#displaytable td,#displaytable th {border: 2px dotted #6b6a6a;} 
            #displaytable table {border-collapse: collapse;width: 100%;}
            #displaytable th{height: 40px;}
            #displaytable td {padding: 6px;}
        </style>
        <script>
            
            function update() {
                var cust_id = document.forms["F12"]["cust_id"].value;
                var attribute = document.forms["F12"]["attribute"].value;
                var changedValue = document.forms["F12"]["changedValue"].value;


                var err = "";
                var op = 1;
                if (attribute == "" || cust_id == "" || changedValue == "") {
                    op = 0;
                    err = err + "Enter all the fields";
                }
                if (attribute == "phone" || attribute == "pincode") {
                    if (/[A-Z]+/.test(changedValue) || /[a-z]+/.test(changedValue)) {
                        op = 0;
                        err = err + ",Changed value should be Numeric if " + attribute + " is selected";
                    }
                }

                if (op == 0) {
                    alert(err);
                }
                else {


                    /*document.F12.method = "POST";
                    document.F12.action = "updateCustomer.asp?cust_id=" + cust_id + "&attribute=" + attribute + "&changedValue=" + changedValue;
                    document.F12.submit();*/
                    var xhttp = new XMLHttpRequest();
                    var op = "";
                    xhttp.onreadystatechange = function () {
                        if (this.readyState == 4 && this.status == 200) {
                            op = this.responseText;
                            if (op != "1") {
                                document.F12.method = "POST";
                                document.F12.action = "selectCustomerupdate.asp?success=y";
                                document.F12.submit();
                            }
                            else if (op == "1") {
                                alert("Records Not Update Retry Updating");
                            }
                        }
                    };
                    xhttp.open("POST", "updateCustomer.asp?cust_id=" + cust_id + "&attribute=" + attribute + "&changedValue=" + changedValue , false);
                    xhttp.send();
                }
            }
        </script>
    </head>
    
        <%
            if session("empid") = "" then 
            response.redirect("sessionTimedOut.asp")
            end if
            Qstr = request.querystring("success")

                Dim Connection
	            Dim Recordset
	            Dim SQL

	            Set Connection = Server.CreateObject("ADODB.Connection")
	            Set Recordset = Server.CreateObject("ADODB.Recordset")

	            Connection.Open "DSN=csidsn"
	            SQL = "SELECT * FROM customer"
	            Recordset.Open SQL,Connection

            content = ""
            if Qstr = "y" then 
                content = content & "<p>Records updated Successfully<br><a href=customer0.asp>Customers</a></p>&nbsp<a href=selectCustomer.asp>Update Records</a>"
            else
                content = content & "Select the following details to update the records </br></br><form name=F12><table><tr><td>Customer ID</td><td><select name=cust_id ><option></option>"
                Do While NOT Recordset.Eof
                content = content & "<option>" & Recordset("cust_id") & "</option>"
                Recordset.MoveNext
                Loop
                
                Recordset.Movefirst
                content = content & "</select></td></tr><tr><td>Attribute</td><td><select name=attribute ><option ></option><option value=fname>First Name</option><option value=lname>Last Name</option><option value=email>Email</option><option value=phone>Phone</option><option value=address>Address</option><option value=city>City</option><option value=state>State</option><option value=pincode>Pincode</option></select></td></tr><tr><td>Enter New Value</td><td><input type=text name=changedValue ></td></tr><tr><td></td><td><input type=button onclick=update() value=Change></td></tr> </table></form>"
                content = content & "</br></br><div id=displaytable><table border = 1><tr><th>Customer ID</th><th>First Name</th><th>Last Name</th><th>Email</th><th>Phone</th><th>Address</th><th>City</th><th>State</th><th>Pincode</th></tr>"
                Do While NOT Recordset.Eof
                    content = content & "<tr><td><center>"&Recordset("cust_id")&"</center></td><td><center>"&Recordset("fname")&"</td><td><center>"&Recordset("lname")&"</td><td><center>"&Recordset("email")&"</td><td><center>"&Recordset("phone")&"</td><td><center>"&Recordset("address")&"</td><td><center>"&Recordset("city")&"</td><td><center>"&Recordset("state")&"</td><td><center>"&Recordset("pincode")&"</td>"
                    Recordset.MoveNext
                Loop
                
                content = content & "</table></div>"
                Recordset.Close
	            Set Recordset=nothing
	            Connection.Close
	            Set Connection=nothing
            end if

        %>
        
    <body>
        <p id="links">
            <a href="customer0.asp">Customer</a>
        </p>
        
       <div id="content">
        <%
            response.write content
        %>    
       </div>
        
        
        
    </body>
</html>
