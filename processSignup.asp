<%@ Language="VBScript" %>

<%
    dim day
    dim month
    dim year
    fname = Request.QueryString("fname")
    lname = Request.QueryString("lname")
    userid = Request.QueryString("userid")
    newpassword = Request.QueryString("newpassword")
    day = Request.QueryString("day")
    month = Request.QueryString("month")
    year = Request.QueryString("year")
    sex = Request.QueryString("sex")
    email = Request.QueryString("email")
    gender = Request.QueryString("sex")
    Qstr= Request.queryString("opr")

    Dim Connection
	Dim Recordset
	Dim SQL

	Set Connection = Server.CreateObject("ADODB.Connection")
    Set Recordset = Server.CreateObject("ADODB.Recordset")
	Connection.Open "DSN=csidsn"

    if Qstr = "fail" then
        Response.write "<head><title>FAIL</title><style>p{position:absolute;left:5%;top:10%;background-color:orange;width:1000px;border-radius:10px;}</style></head><p style=font-size:40px;color:red;>Username/email already exists Please enter another username/email<br><a href=index.html>Sign Up</a>&nbsp&nbsp<a href = forgotpassword.asp>Forgot Passowrd?</a></p>"
    elseif Qstr = "signup" then
        f=0
        SQL = "select * from employee"
        Recordset.Open SQL,Connection
        Do While NOT Recordset.Eof
            if Recordset("userid") = userid then
                f=1
            else
                f=0
            end if

            if Recordset("email") = email then
                f=1
            else
                f=0
            end if

            if f = 1 then 
                exit do 
            else
	            Recordset.MoveNext
            end if
        loop
        
        Set Recordset=nothing
    
        if f =1 then
            response.redirect("processSignup.asp?opr=fail")
        else


            Set Recordset = Server.CreateObject("ADODB.Recordset")
	        SQL = "insert into employee (fname,lname,email,userid,password,dob,gender) values ('"& fname &"','"&lname&"','"&email&"','"&userid&"','"&newpassword&"','"&year&"-"&month&"-"&day&"','"&gender&"');"
	
	        Connection.Execute SQL,recaffected
            if err<>0 then
                response.write "No permissions to write"
            else
                response.write "<head><title>Success</title><style>p{position:absolute;left:5%;top:10%;background-color:orange;width:1000px;border-radius:10px;}</style></head><p style=font-size:50px;color:rgba(0, 255, 33, 0.90);>Sign Up Successfull! <a href=index.html>Log In</a></p>"
                
            end if

    
	        Set Recordset=nothing
	
	    
        end if
        
    end if
    Set Connection=nothing

%>
