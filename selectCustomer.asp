<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
        <script>
            function getCustId(id) {
                var custid = document.getElementById(id).value;
                document.getElementById("cust_id").value = custid;
            }
            function transfer() {
                var cust_id = document.getElementById("cust_id").value;

                var f = 0;
                if (cust_id == "")
                    f = 1;
                /*document.F3.method = "POST";
                document.F3.action = "storeCookie.asp?cust_id="+cust_id;
                document.F3.submit();*/

                if (f == 1) {
                    alert("Select Customer Name");
                    return false;
                }
                else {
                    var xhttp = new XMLHttpRequest();
                    var op = "";
                    xhttp.onreadystatechange = function () {
                        if (this.readyState == 4 && this.status == 200) {
                            document.F3.metod = "Post";
                            document.F3.action = "order.asp";
                            document.F3.submit();

                        }
                    };
                    xhttp.open("POST", "storeCookie.asp?cust_id=" + cust_id, false);
                    xhttp.send();
                }
            }
        </script>
        <style>
            #content
            {
                position: absolute;
                top: 25%;
                left: 12%;
                background-color: rgba(128, 128, 128, 0.77);
                padding: 30px;
                width: 700px;
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
	    Set Recordset = Server.CreateObject("ADODB.Recordset")
                    

	    Connection.Open "DSN=csidsn"
	    SQL = "SELECT * FROM customer;"
	    Recordset.Open SQL,Connection
                    
    %>
    <body>
        <form name="F3">
        <div id="content">
            <fieldset>
                <legend>Customer Details</legend>
                <table >
                    <tr>
                        <td>Customer Name</td>
                        <td></td>
                        <td>
                                
                            <select id="custname" onchange="getCustId(this.id)">
                                <%
                                    response.write "<option></option>"
                                    Do While NOT Recordset.Eof
                                        response.write "<option value="&Recordset("cust_id")&">"&Recordset("fname")&" "&Recordset("lname")&"</option>"
                                        Recordset.MoveNext
                                    loop
                                %>
                            </select>
                                
                        </td>
                        <td></td>
                        <td></td>
                        <td>Customer Id</td>
                        <td></td>
                        <td> <input type="text" disabled id="cust_id" size="15"></td>
                    </tr>
                    </table></br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
            <input type="button" style ="background: #00ff21;"value ="Proceed" onclick="return transfer();">
            <br>
        </fieldset>
            </div>
        </br>        
        </form>
    </body>
</html>
