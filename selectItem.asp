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
            a
            {
                text-decoration: none;
            }
            p a
            {
                text-decoration: none;
            }
            #displaytable table,#displaytable td,#displaytable th {border: 2px dotted #6b6a6a;} 
            #displaytable table {border-collapse: collapse;width: 100%;}
            #displaytable th{height: 40px;}
            #displaytable td {padding: 6px;}
            
        </style>
        <script>

            function update() {
                var itemno = document.forms["F12"]["itemno"].value;
                var attribute = document.forms["F12"]["attribute"].value;
                var changedValue = document.forms["F12"]["changedValue"].value;


                var err = "";
                var op = 1;
                if (attribute == "" || itemno == "" || changedValue == "") {
                    op = 0;
                    err = err + "Enter all th fields";
                }
                if (attribute == "stock_quantity" || attribute == "cost_per_item") {
                    if (/[A-Z]+/.test(changedValue) || /[a-z]+/.test(changedValue)) {
                        op = 0;
                        err = err + ",Changed value should be Numeric if " + attribute + " is selected";
                    }
                }
                if (attribute == "item_discontinued") {
                    if (!(changedValue == "y" || changedValue == "Y" || changedValue == "N" || changedValue == "n")) {
                        op = 0;
                        err = err + ", New Value should be 'Y' or 'N' ";
                    }
                }

                if (op == 0) {
                    alert(err);
                }
                else {


                    /*document.F12.method = "POST";
                    document.F12.action = "updateItem.asp?itemno=" + itemno + "&attribute=" + attribute + "&changedValue=" + changedValue;
                    document.F12.submit();*/
                    var xhttp = new XMLHttpRequest();
                    var op = "";
                    xhttp.onreadystatechange = function () {
                        if (this.readyState == 4 && this.status == 200) {
                            op = this.responseText;
                            if (op != "1") {
                                
                                document.F12.action = "selectItem.asp?success=y";
                                document.F12.submit();
                            }
                            else if (op == "1") {
                                alert("Records Not Update Retry Updating");
                            }
                        }
                    };
                    xhttp.open("POST", "updateItem.asp?itemno=" + itemno + "&attribute=" + attribute + "&changedValue=" + changedValue, false);
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
	            SQL = "SELECT * FROM inventory where item_discontinued = 'n'"
	            Recordset.Open SQL,Connection

            content = ""
            if Qstr = "y" then 
                content = content & "<p>Records updated Successfully<br><a href=Inventory.asp>Inventory</a></p>&nbsp<a href=selectItem.asp>Update Records</a>"
            else
                content1 = content1 & "<p id=links><a href=Inventory.asp>Inventory</a></p>"
                content = content & "Select the following details to update the records </br></br><form name=F12 method = Post><table><tr><td>Item Number</td><td><select name=itemno ><option></option>"
                Do While NOT Recordset.Eof
                    content = content & "<option>" & Recordset("itemno") & "</option>"
                    Recordset.MoveNext
                Loop
                
                Recordset.Movefirst
                content = content & "</select></td></tr><tr><td>Attribute</td><td><select name=attribute ><option ></option><option value=item_name>Item Name</option><option value=describtion>Describtion</option><option value=cost_per_item>Cost per Item</option><option value=stock_quantity>Stock Quantity</option><option value = item_discontinued>Item Discontinued (Y/N)</option></select></td></tr><tr><td>Enter New Value</td><td><input type=text name=changedValue ></td></tr><tr><td></td><td><input type=button onclick=update() value=Change></td></tr> </table></form>"
                content = content & "</br></br><div id='displaytable'><table ><tr><th>Item No</th><th>Date of last order</th><th>Item Name</th><th>Stock Location</th><th>Describtion</th><th>Cost per Item</th><th>Stock Quantity</th><th>Total Value</th><th>Item Discontinued?</th></tr>"
                Do While NOT Recordset.Eof
                    if Recordset("stock_quantity") = 0 then
                        content = content & "<tr bgcolor=red><td><center>"&Recordset("itemno")&"</center></td><td><center>"&Recordset("date_of_last_order")&"</td><td><center>"&Recordset("item_name")&"</td><td><center>"&Recordset("stock_location")&"</td><td><center>"&Recordset("describtion")&"</td><td><center>"&Recordset("cost_per_item")&"</td><td><center>"&Recordset("stock_quantity")&"</td><td><center>"&Recordset("total_value")&"</td>"
                        if Recordset("item_discontinued") = "y" then
                            content = content & "<td><center>Yes</td></tr>"
                        else
                            content = content & "<td><center>No</td></tr>"
                        end if
                    elseif Recordset("stock_quantity") < 6 then
                        content = content & "<tr bgcolor=#ff00dc><td><center>"&Recordset("itemno")&"</center></td><td><center>"&Recordset("date_of_last_order")&"</td><td><center>"&Recordset("item_name")&"</td><td><center>"&Recordset("stock_location")&"</td><td><center>"&Recordset("describtion")&"</td><td><center>"&Recordset("cost_per_item")&"</td><td><center>"&Recordset("stock_quantity")&"</td><td><center>"&Recordset("total_value")&"</td>"
                        if Recordset("item_discontinued") = "y" then
                            content = content & "<td><center>Yes</td></tr>"
                        else
                            content = content & "<td><center>No</td></tr>"
                        end if
                    else
                        content = content & "<tr><td><center>"&Recordset("itemno")&"</center></td><td><center>"&Recordset("date_of_last_order")&"</td><td><center>"&Recordset("item_name")&"</td><td><center>"&Recordset("stock_location")&"</td><td><center>"&Recordset("describtion")&"</td><td><center>"&Recordset("cost_per_item")&"</td><td><center>"&Recordset("stock_quantity")&"</td><td><center>"&Recordset("total_value")&"</td>"
                        if Recordset("item_discontinued") = "y" then
                            content = content & "<td><center>Yes</td></tr>"
                        else
                            content = content & "<td><center>No</td></tr>"
                        end if
                    end if
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
        
        <div>
            <%
                response.write content1
            %>
        </div>
       <div id="content">
        <%
            response.write content
        %>    
       </div>
        
        
        
    </body>
</html>
