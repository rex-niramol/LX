<%
Session.Contents.RemoveAll()
'Session("promark_login") = ""
url2 = "index.asp"
response.Redirect(url2)
%>