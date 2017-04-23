<%@ Language="VBScript" %>
<%
	uname = Request.QueryString("uname")
	pname = Request.QueryString("pname")

	Dim Connection
	Dim Recordset
	Dim SQL

	Set Connection = Server.CreateObject("ADODB.Connection")
	Set Recordset = Server.CreateObject("ADODB.Recordset")

	Connection.Open "DSN=csidsn"
	SQL = "SELECT * FROM employee;"
	
	Recordset.Open SQL,Connection
	
	f = 0
    empid = 0
	Do While NOT Recordset.Eof
        empid = 0
		If uname = Recordset("userid") Then 
			f = 0
            If pname = Recordset("password") Then
			    f = 0
                empid = Recordset("empid")
		    Else
			    f = 1
		    end if
		Else       
			f = 1
        end if

        if f = 0 then
            exit do 
        else
	        Recordset.MoveNext
        end if
	Loop
	
	If f = 1 then
		response.write "1"
	else
		response.write empid
        session("empid") = empid
        
	end if
	
	Recordset.Close
	Set Recordset=nothing
	Connection.Close
	Set Connection=nothing
%>
