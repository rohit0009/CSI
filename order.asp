<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Order</title>
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
            
            #content
            {
                overflow: hidden;
                position: absolute;
                top: 15%;
                left: 13%;
                width: 60%;
                height: auto;
                background-color: rgba(128, 128, 128, 0.77);
                padding: 60px;
            }
            
            
            
            

            
            
            
        </style>
        <script>

            function getItemNo(id) {
                var itemno = document.getElementById(id).value;
                document.getElementById("itemno").value = itemno;
                /*document.F3.method = "POST";
                document.F3.action = "getFromDb.asp?itemno=" + itemno;
                document.F3.submit();*/
                var xhttp = new XMLHttpRequest();
                var op = "";
                xhttp.onreadystatechange = function () {
                    if (this.readyState == 4 && this.status == 200) {
                        op = this.responseText;
                        var result = op.split(" ");

                        if (op != "") {
                            document.getElementById("cost_per_item").value = result[0];
                            document.getElementById("stock_quantity").value = result[1];
                        }
                        else if (op == "") {
                            document.getElementById("cost_per_item").value = "";
                        }
                    }
                };
                xhttp.open("POST", "getFromDb.asp?itemno=" + itemno, false);
                xhttp.send();
            }
            function calculate() {
                var itemno = document.getElementById("itemno").value;
                var stock_quantity = parseInt(document.getElementById("stock_quantity").value);
                var cost_per_item = parseInt(document.getElementById("cost_per_item").value);
                
                var ordered_quantity = parseInt(document.getElementById("ordered_quantity").value);

                
                
                var f = 0;
                var err = "";

                if (ordered_quantity == "") {
                    f = 1;
                    err = err + "Enter Quantity";
                }
                if (itemno == "") {
                    f = 1
                    err = err + " ,Select Item";
                }
                if ((/[A-Z]+/.test(ordered_quantity)) || (/[a-z]+/.test(ordered_quantity))) {
                    f = 1;
                    err = err + "Quantity should the numeric";
                }


                if (ordered_quantity > stock_quantity) {
                    f = 1;
                    if (err == "")
                        err = err + "This quantity cannot be ordered";
                    else
                        err = err + " ,This quantity cannot be ordered";
                }

                if (f == 1)
                    alert(err);
                else     {
                    /*document.F3.method = "POST";
                    document.F3.action = "insertIntoMyCart.asp?itemno=" + itemno + "&ordered_quantity=" + ordered_quantity+"&cost_per_item="+cost_per_item;
                    document.F3.submit();*/
                    var xhttp = new XMLHttpRequest();
                    var op = "";
                    xhttp.onreadystatechange = function () {
                        if (this.readyState == 4 && this.status == 200) {
                            op = this.responseText;

                            if (op != "1") {
                                document.F3.method = "Post";
                                document.F3.action = "showCart.asp";
                                document.F3.submit();
                            }
                            else if (op == "1") {
                                alert("Error Inserting into cart Session expired");
                            }
                        }
                    };
                    xhttp.open("POST", "insertIntoMyCart.asp?itemno=" + itemno + "&ordered_quantity=" + ordered_quantity+"&cost_per_item="+cost_per_item, false);
                    xhttp.send();    
                }
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
	                
                    Set RecordsetItem = Server.CreateObject("ADODB.Recordset")

	                Connection.Open "DSN=csidsn"
	                
                    SQL = ""
                    SQL = "select * from inventory where item_discontinued = 'n';"
                    RecordsetItem.Open SQL,Connection

    %>
    <body>
    
            <div id="content">
                <form name="F3">
                    
                    <fieldset>
                        <legend>Order Details</legend></br>
                        
                        &nbsp&nbsp&nbsp&nbsp&nbspItem Name &nbsp&nbsp&nbsp&nbsp
                        <select id="item_name" onchange="getItemNo(this.id)">
                            <option></option>
                            <%
                                Do While NOT RecordsetItem.Eof
                                    if RecordsetItem("stock_quantity") = 0 then
                                        response.write "<option value="&RecordsetItem("itemno")&" style='background-color:red;'>"&RecordsetItem("item_name")&"</option>"
                                    else
                                        response.write "<option value="&RecordsetItem("itemno")&">"&RecordsetItem("item_name")&"</option>"
                                    end if
                                    RecordsetItem.MoveNext
                                loop
                            %>
                        </select>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                        Available Stock &nbsp&nbsp&nbsp&nbsp <input type="text" disabled id="stock_quantity" value="0" size="5" >
                        </br></br></br>&nbsp&nbsp&nbsp&nbsp&nbsp
                        Item Number 
                        <input type="text" disabled id="itemno" size="5"> </input>&nbsp&nbsp&nbsp&nbsp&nbsp
                        Cost Per Item &nbsp&nbsp&nbsp&nbsp <input type="text" disabled size="5" id="cost_per_item"> /- &nbsp&nbsp&nbsp&nbsp&nbsp
                        Quantity &nbsp&nbsp&nbsp&nbsp <select id="ordered_quantity"><option></option><option value=1>1</option><option value=2>2</option><option value=3>3</option><option value=4>4</option><option value=5>5</option><option value=6>6</option><option value=7>7</option><option value=8>8</option><option value=9>9</option><option value=10>10</option></select> &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                        </br></br></br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
                        <input type="button" value="Add Order" style="background-color: #00ff21" onclick="calculate();"></br></br></br>
                        
                    </fieldset>
                </form>

                <%
                    Set RecordsetItem=nothing                    
	                Set Connection=nothing
                %>
            </div>
        </div>
    </body>
</html>
