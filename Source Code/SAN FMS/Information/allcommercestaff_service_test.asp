<!--#Include virtual="/Services/LDAP_Login/service_functions.asp" -->
<%

str_group = "Staff_G"

 	response.redirect "http://www.commerce.uct.ac.za/Services/LDAP_Login/allcommercestaff_service.php?group=" & str_group
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 1</title>
</head>

<body>
<h1>All Commerce Staff Test</h1>
</body>

</html>
