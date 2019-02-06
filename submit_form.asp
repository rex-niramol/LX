<!--#include file ="conf/inc_config.asp"-->
<%

username = str_filter(request.form("username"))
password = en_pass(request.form("password"))

if username <> "" and password <> "" Then

 	status="NOT_success"

	strSQL = "SELECT top 1 username, firstname, lastname FROM hmc_tb_m_users where username = '"& username &"' and password = '"& password &"' and active_sts ='1'"

	Set objRec = Server.CreateObject("ADODB.Recordset")
	objRec.Open strSQL, Conn, 1,3
		
	If Not objRec.EOF then
		
		Session("hmc_login") = True
        Session("hmc_username") = objRec("username")
        Session("hmc_firstname") = objRec("firstname")
        Session("hmc_lastname") = objRec("lastname")
        Response.write "HMC"

        '***Insert to Log table***'
            If Request.ServerVariables("HTTP_CLIENT_IP") <> ""  then
                varIPcreatorx=Request.ServerVariables("HTTP_CLIENT_IP")
            ElseIf  Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "" then
                varIPcreatorx=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
            ElseIf  Request.ServerVariables("remote_addr") <> ""  then
                varIPcreatorx=Request.ServerVariables("remote_addr")
            Else
                varIPcreatorx=Session("hmc_username")
            End If

            cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" loged-in', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
            Set objExec = Conn.Execute(cmdSql)
            Set objExec = Nothing
        '***Insert to Log table***'
		
	Else 
	    Session("hmc_login") = False
	    Response.write "NO"

        '***Insert to Log table***'
            If Request.ServerVariables("HTTP_CLIENT_IP") <> ""  then
                varIPcreatorx=Request.ServerVariables("HTTP_CLIENT_IP")
            ElseIf  Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "" then
                varIPcreatorx=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
            ElseIf  Request.ServerVariables("remote_addr") <> ""  then
                varIPcreatorx=Request.ServerVariables("remote_addr")
            Else
                varIPcreatorx=Session("hmc_username")
            End If

            cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('log-in failed !', '0', N'"&varIPcreatorx&"', N'"&varIPcreatorx&"'); "
            Set objExec = Conn.Execute(cmdSql)
            Set objExec = Nothing
        '***Insert to Log table***'
	End If

	objRec.Close()
	Set objRec = Nothing
	
	
Else
    Response.write "NO"
    Session("authenmessage") = "Username OR Password incorrect!"    
    Session("hmc_login") = False

End If
%> 



