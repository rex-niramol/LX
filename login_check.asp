<!--#include file ="conf/inc_config.asp"-->
<%

cardnumber = Request.Form("txt_cardnumber")
datecheck = request.form("date_check_in")


If cardnumber <> "" then
Session("guest_number") = cardnumber
guestnumber = Session("guest_number")
End If


if datecheck <> "" and guestnumber <> "" Then


 	status="NOT_success"
	'Session("promark_login") = False
	cardnumber = request.form("txt_cardnumber")
	date_check_in = request.form("date_check_in")
	exp_date = date_check_in
	'Split Date Format
    d = split(exp_date,"/")
    exp_date = d(2) & "-" & d(1) & "-" & d(0)


	strSQL = "SELECT top 1 guest_id FROM pm_tb_r_sales where guest_id = '"& cardnumber &"' and CAST(exp_dt AS date) = '"& exp_date &"' and active_sts ='1' order by updated_dt"

	Set objRec = Server.CreateObject("ADODB.Recordset")
	objRec.Open strSQL, Conn, 1,3
		
	If Not objRec.EOF then
		
		guest_id = objRec("guest_id")
		
		status = "success"
		
		If status = "success" then
		
		Session("promark_login") = True
		

		Response.write "<form method='post' name = 'voucher' action='voucher.asp'>"
		Response.write "<input type='hidden' id='guest_id' name='guest_id' value='"&guest_id&"'>"
		Response.write "<input type='submit' style='display: none;'>"
		Response.write "</form>"
		
		Response.write "<script type='text/javascript'>"
  		Response.write "document.forms['voucher'].submit();"
		response.write "</script>"
		
		End If
		'response.Redirect(url)

	'objRec.MoveNext
	'Wend
	
	Else 
	cardnumber = ""
	date_check_in = ""
	'Session("promark_login") = False
	url2 = "index.asp"
	response.Redirect(url2)
	
	End If
	objRec.Close()
	Set objRec = Nothing
	
	
Else

url2 = "index.asp"
response.Redirect url2
Session("userlogin") = False
End If
%> 



