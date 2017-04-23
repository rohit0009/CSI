<%@ Language="VBScript" %>
<%
    if session("empid") = "" then 
                    response.redirect("sessionTimedOut.asp")
                    end if
    cust_id = request.querystring("cust_id")
    session("cust_id") = cust_id
    response.write session("cust_id")
    
%>