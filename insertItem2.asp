<!DOCTYPE html>

<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title></title>
        <script language="JavaScript">
            function validate() {
                var iname = document.forms["F1"]["iname"].value;
                
                var stockL = document.forms["F1"]["stockL"].value;
                var desc = document.forms["F1"]["desc"].value;
                var cost = document.forms["F1"]["cost"].value;
                var qty = document.forms["F1"]["qty"].value;

                var err = "";
                var f = 0;

                if (iname == "" || vname == "" || stockL == "" || desc == "" || cost == "" || qty == "") {
                    f = 1;
                    err = err + "Please fill all the above fields";
                }

                else {
                    if (!(/[0-9]+/.test(cost))) {
                        f = 1;
                        err = err + "Please enter a numeric value for cost of item. "
                    }

                    if (!(/[0-9]+/.test(qty))) {
                        f = 1;
                        err = err + "Please enter a numeric value for stock quantity"
                    }
                }

                if (f == 1) {
                    alert(err);
                    return false;
                }
                else {
                    return true;
                }
            }
        </script>
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
        </style>
    </head>
    <body>
        <center>
        <div id="content">
        <form name="F1" method="post" action="insertItem.asp">
            
            <table>
                <tr>
                    <td>
                        Enter item name : &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<input type="text" name="iname" class="textfield1" placeholder="Item Name">
                    </td>
                </tr>
                
                <tr>
                    <td>
                        Enter stock location : &nbsp&nbsp&nbsp<input type="text" name="stockL" class="textfield3" placeholder="Stock Location">
                    </td>
                </tr>
                <tr>
                    <td>
                        Description : &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<input type="text" name="desc" class="textfield3" placeholder="Description">
                    </td>
                </tr>
                <tr>
                    <td>
                        Cost per item : &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<input type="text" name="cost" class="textfield4" placeholder="Cost per item">
                    </td>
                </tr>
                <tr>
                    <td>
                        Stock quantity : &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<input type="text" name="qty" class="textfield5" placeholder="Stock Quantity">
                    </td>
                </tr>
                <tr>
                    <td>
                        &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<input type="submit" value="Add to Inventory" onclick="return validate()">
                    </td>
                </tr>
            </table>
           
            </form>
        </div>
        </center>
    </body>
</html>
