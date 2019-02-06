<!--#include file ="conf/inc_config.asp"-->
<%
	If session("admin_lang_code") = "TH" Then
        varLangSend="th/"
	Else
        varLangSend=""
	End If

    offer_code="hmc01"
    card_code="HMCCOM"
    prop_code = "CHR"

    prop_choose = request.Form("property")
    roomtype = request.Form("roomtype")
    date_arrival = request.Form("arrival")
        d = split(date_arrival,"/")
        varArrDate = d(2) & "-" & d(1) & "-" & d(0)
    date_departure = request.Form("departure")
        d = split(date_departure,"/")
        varDptDate = d(2) & "-" & d(1) & "-" & d(0)
    adult = request.Form("adult")
    child = request.Form("child")
    title = request.Form("title")
    firstname = request.Form("firstname")
    lastname = request.Form("lastname")
    email = request.Form("email")
    phone = request.Form("phone")

    crd01 = request.Form("crd01")
    crd02 = request.Form("crd02")
    crd03 = request.Form("crd03")
    crd04 = request.Form("crd04")

    creditno = crd01&crd02&crd03&crd04
        creditFirstPart = Mid(creditno, 1, 6)
        creditMiddlePart = Mid(creditno, 7, 6)
        creditLastPart = Mid(creditno, 13, 4)
        creditPms=creditFirstPart&"xxxxxx"&creditLastPart
        creditStore="xxxxxx"&creditMiddlePart&"xxxx"
    expCreditMM = request.Form("expCreditMM")
    expCreditYY = request.Form("expCreditYY")
        expCredit=expCreditMM&"/"&expCreditYY
        expCreditPMS="xx/"&Right(String(2, "0") & expCreditYY, 2)

    'Response.write email

    '**Add More**'
    cust_id = request.Form("membership_no")
    cust_expiry = request.Form("expiry_date")
        d = split(cust_expiry,"/")
        varExpDate = d(2) & "-" & d(1) & "-" & d(0)
    redeem_type=request.Form("redeem_type")
    mobile_code=request.Form("mobile_code")
    hmc_code=request.Form("hmc_code")

    varAllot=2
    pmsSystem="COMANCHE"
    pmsSystemName="COMANCHE"
    remark=request.Form("remark")&" (HMC, MEM: "&cust_id&", EXP: "&varExpDate&", REDT: "&redeem_type&", REDC: "&mobile_code&", RESC: "&hmc_code&")"

    varRecPriceAgreed="0"
    varRecPriceShow="0"
    varBookingSts="C"
    varCardName=firstname&" "&lastname

    varStatus="success"

    If Request.ServerVariables("HTTP_CLIENT_IP") <> ""  then
        varIPcreatorx=Request.ServerVariables("HTTP_CLIENT_IP")
    ElseIf  Request.ServerVariables("HTTP_X_FORWARDED_FOR") <> "" then
        varIPcreatorx=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    ElseIf  Request.ServerVariables("remote_addr") <> ""  then
        varIPcreatorx=Request.ServerVariables("remote_addr")
    Else
        varIPcreatorx=Session("hmc_username")
    End If

'    '******Allotment******'
    '    arrArrival=Split(date_arrival,"/")
    '    dayArrival=arrArrival(0)
    '    cmdSql = "SELECT (allot.allotment-b.allotment) as 'allot_no' from pm_tb_a_allotment allot left join pm_tb_a_allotment b  on(allot.prop_code=b.prop_code) where allot.active_sts='1' and allot.prop_code='"&prop_code&"' and allot.day='1' and b.day='"&CInt(dayArrival)&"' "
    '    xxx=cmdSql
    '    Set recSql = Server.CreateObject("ADODB.Recordset")
    '    recSql.Open cmdSql, Conn, 1,3
    '
    '    If Not recSql.EOF then
    '        varAllot=2'CInt(recSql("allot_no"))
    '    End If
    '    recSql.Close()
    '    Set recSql = Nothing
'    '******Allotment******'

    '******Check Exist******'
    varGuestID=""
        cmdSql = " SELECT top 1 lower(guest_id) as 'guest_id' from hmc_tb_r_booking where ltrim(rtrim(lower(guest_id)))='"&LCase(Trim(cust_id))&"' and active_sts='1' and delete_sts='N'"

        Set recSql = Server.CreateObject("ADODB.Recordset")
        recSql.Open cmdSql, Conn, 1,3

        While Not recSql.EOF
            varGuestID=recSql("guest_id")
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

            cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" try to add. The booking already exist!', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
            Set objExec = Conn.Execute(cmdSql)
            Set objExec = Nothing
        '***Insert to Log table***'
        recSql.MoveNext
        Wend
        recSql.Close()
        Set recSql = Nothing
    '******Check Exist******'

    If varAllot>0 and varGuestID="" Then

        '******Pricing******'
            cmdSql = " SELECT offer_price_show, offer_price_agreed, hotels_availability from hmc_tb_a_offers_pricing where active_sts='1' "

            Set recSql = Server.CreateObject("ADODB.Recordset")
            recSql.Open cmdSql, Conn, 1,3

            While Not recSql.EOF
                varPriceShow=recSql("offer_price_show")
                varPriceAgreed=recSql("offer_price_agreed")
                varHotels=recSql("hotels_availability")

                '******Hotel Selected******'
                arrHotels=Split(varHotels,",")
                If IsArray(arrHotels) Then
                    for each varVal in arrHotels
                        If varVal=prop_choose Then
                            varRecPriceAgreed=varPriceAgreed
                            varRecPriceShow=varPriceShow
                            Exit For
                        End If
                    Next
                End If
                '******Hotel Selected******'
            recSql.MoveNext
            Wend
            recSql.Close()
            Set recSql = Nothing
        '******Pricing******'

        '******Extra Info*****'
        market_segment="LOYAL"
        '******Extra Info*****'

        '******Property Info******'
            cmdSql = " SELECT email_resv,prop_name from properties where prop_code='"&prop_choose&"' "

            Set recSql = Server.CreateObject("ADODB.Recordset")
            recSql.Open cmdSql, Conn, 1,3

            While Not recSql.EOF
                emailSendToResv=recSql("email_resv")
                emailPropName=recSql("prop_name")
            recSql.MoveNext
            Wend
            recSql.Close()
            Set recSql = Nothing

            'extra request
            If LCase(prop_choose)="cgcw" Then
                emailSendToResv="cgcwreservation@chr.co.th"
            ElseIf LCase(prop_choose)="cgc" Then
                emailSendToResv="rsvncgc@chr.co.th"
            ElseIf LCase(prop_choose)="czsc" Then
                emailSendToResv="czsc@cosihotels.com"
            End If
        '******Property Info******'

        '******PMSsystem******'
            cmdSql = " SELECT sys_code, sys_name, hotels_avai from hmc_tb_m_pms where active_sts='1' "

            Set recSql = Server.CreateObject("ADODB.Recordset")
            recSql.Open cmdSql, Conn, 1,3

            While Not recSql.EOF
                sys_code=recSql("sys_code")
                sys_name=recSql("sys_name")
                varHotels=recSql("hotels_avai")

                '******Hotel Selected******'
                arrHotels=Split(varHotels,",")
                If IsArray(arrHotels) Then
                    for each varVal in arrHotels
                        If varVal=prop_choose Then
                            pmsSystem = sys_code
                            pmsSystemName=sys_name
                            Exit for
                        End If
                    Next
                End If
                '******Hotel Selected******'

            recSql.MoveNext
            Wend
            recSql.Close()
            Set recSql = Nothing
        '******PMSsystem******'


'        '******CommandChe Web Service
'            Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
'            oXmlHTTP.Open "POST", "http://www.chrnetwork.com/promark/ProMark.asmx", False 
'
'            oXmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8" 
'            oXmlHTTP.setRequestHeader "SOAPAction", "http://tempuri.org/saveBooking"
'
'            SOAPRequest ="<?xml version=""1.0"" encoding=""utf-8""?>" & _
'              "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
'              "<soap:Body>" & _
'                "<saveBooking xmlns=""http://tempuri.org/"">" &_
'                      "<property>"&prop_choose&"</property>" &_
'                      "<title>"&title&"</title>" &_
'                      "<first_name>"&firstname&"</first_name>" &_
'                      "<last_name>"&lastname&"</last_name>" &_
'                      "<email>"&email&"</email>" &_
'                      "<phone>"&phone&"</phone>" &_
'                      "<arrive>"&date_arrival&"</arrive>" &_
'                      "<depart>"&date_departure&"</depart>" &_
'                      "<room_night>1</room_night>" &_
'                      "<room_type>"&roomtype&"</room_type>" &_
'                      "<adult>"&adult&"</adult>" &_
'                      "<child>"&child&"</child>" &_
'                      "<credit_no>"&creditPms&"</credit_no>" &_
'                      "<price>"&varRecPriceAgreed&"</price>" &_
'                      "<business_code>"&offer_code&"</business_code>" &_
'                      "<guest_id>"&cust_id&"</guest_id>" &_
'                      "<pms>"&pmsSystem&"</pms>" &_
'                      "<card_exp>"&expCreditPMS&"</card_exp>" &_
'                      "<card_type>"&card_code&"</card_type>" &_
'                      "<market>"&market_segment&"</market>" &_
'                      "<remark>"&remark&"</remark>" &_
'                      "<card_owner>"&UCase(prop_code)&"</card_owner>" &_
'                "</saveBooking>" &_
'              "</soap:Body>" &_
'              "</soap:Envelope>"
'
'            oXmlHTTP.send SOAPRequest  
'            varResponse= oXmlHTTP.responseText
'            function findStringBetween(s_string, s_pre, s_post, n_options)
'              ' n_options:
'              '  1: IGNORE case in the search
'              dim a, b
'              ' init
'              findStringBetween = ""
'              ' find pre-word
'              a = instr(1, s_string, s_pre, n_options)
'              if a>0 then
'                a = a + len(s_pre)
'                b = instr(a, s_string, s_post, n_options)
'                if b>0 and not (s_post = "") then
'                  findStringBetween = mid(s_string, a, b-a)
'                else
'                  findStringBetween = mid(s_string, a)
'                end if
'              end If
'            end function  
'
'            varBookingNo = findStringBetween(varResponse, "<reservno>", "</reservno>", 0)
'            varStatus = findStringBetween(varResponse, "<status>", "</status>", 0)

'            '***Insert to Log table***'
'
'                cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" insert to PMS completed.', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
'                Set objExec = Conn.Execute(cmdSql)
'                Set objExec = Nothing
'            '***Insert to Log table***'

'        '******CommandChe Web Service

        '******Inseret to DB******'
            varSuccess=InStr(LCase(varResponse),"success")
            If LCase(varStatus)="success" Then 'Success
                If UCase(pmsSystemName)<>"COMANCHE" Then
                     varBookingSts="P"
                End If

                '******Insert Booking******'

                cmdSql = "INSERT INTO [dbo].[hmc_tb_r_booking]([guest_id], [expiry_dt], [redemption_type], [mobile_code], [hmc_code],  [conf_id], [prop_code], [prop_system], [offer_code], [offer_price_agreed], [offer_price_show], [card_type], [arrival_dt], [departure_dt], [roomtype],	[adult], [child], [creditno], [creditexp], [lang_ref], [email], [card_name], [title], [first_name], [last_name], [primary_phone_no], [note], [sale_sts], [active_sts], [delete_sts], [value_seq], [parent_id], [created_by], [updated_by]) values('"&cust_id&"','"&varExpDate&"','"&redeem_type&"','"&mobile_code&"','"&hmc_code&"', '"&varBookingNo&"', '"&prop_choose&"', '"&pmsSystemName&"', '"&offer_code&"', '"&varRecPriceAgreed&"', '"&varRecPriceShow&"', '"&card_code&"', '"&varArrDate&"', '"&varDptDate&"', '"&roomtype&"', '"&adult&"', '"&child&"', '"&creditStore&"', '"&expCredit&"', 'EN', N'"&email&"', N'"&varCardName&"', N'"&title&"', N'"&firstname&"', N'"&lastname&"', '"&phone&"', N'"&remark&"', '"&varBookingSts&"', '1', 'N', '99', '"&prop_code&"', N'"&Session("hmc_username")&"', N'"&Session("hmc_username")&"'); "
                stringsql=cmdSql
                rpID = "1"
                Response.write cmdSql

'                cmdSql = cmdSql & "Set Nocount no;"
'                cmdSql = cmdSql & "SELECT NewsID = SCOPE_IDENTITY();"
'                cmdSql = cmdSql & "SELECT SCOPE_IDENTITY() AS NewsID;"
'                cmdSql = cmdSql & "SELECT @@IDENTITY AS 'NewsID';"
'                cmdSql = cmdSql & "Set Nocount off;"

                Set objExec = Conn.Execute(cmdSql)
                Set objExec = Nothing
                If Err.Number = 0 Then
                    cmdSql = "SELECT MAX(rn_no) AS NewsID from hmc_tb_r_booking;"
                    Set recNewsID = Server.CreateObject("ADODB.Recordset")
                    recNewsID.Open cmdSql, Conn, 1,3
                    If Not recNewsID.EOF then
                        rpID = recNewsID("NewsID")
                    End If
                    recNewsID.Close()
                    Set recNewsID = Nothing
                    'cmdSQL="update pm_tb_r_booking set conf_id='"&varBookingNo&"' where rn_no='"&rpID&"' "
                    'conn.Execute cmdSQL

                    '***Insert to Log table***'
                        cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" insert rn="&rpID&" completed.', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
                        Set objExec = Conn.Execute(cmdSql)
                        Set objExec = Nothing
                    '***Insert to Log table***'

                    '************SEND EMAIL*********'
                        If UCase(pmsSystem)="COMANCHE" Then
                            pageSuccess="https://www.centarahotelsresorts.com/"&varLangSend&"hmc/pms-success-confirmed.asp?rp="&rpID
                            sendMailToCust="https://www.centarahotelsresorts.com/"&varLangSend&"hmc/inc_file/sendemail-customer-pms-mhtml-comp-confirmed.asp?rp="&rpID
                            sendMailToProp="https://www.centarahotelsresorts.com/"&varLangSend&"hmc/inc_file/sendemail-property-pms-mhtml-comp-confirmed.asp?rp="&rpID
                            sendMailSubject=varBookingNo & " - Booking confirmation for " & emailPropName
                        Else
                            pageSuccess="https://www.centarahotelsresorts.com/"&varLangSend&"hmc/pms-success-noconfirmed.asp?rp="&rpID
                            sendMailToCust="https://www.centarahotelsresorts.com/"&varLangSend&"hmc/inc_file/sendemail-customer-pms-mhtml-comp-noconfirmed.asp?rp="&rpID
                            sendMailToProp="https://www.centarahotelsresorts.com/"&varLangSend&"hmc/inc_file/sendemail-property-pms-mhtml-comp-noconfirmed.asp?rp="&rpID
                            sendMailSubject="Centara Hotels & Resorts - Booking confirmation for " & emailPropName
                        End If

                        '******EMAIL VARIABLE******'
                        Dim iConf, Flds, iMsg
                        Dim Sender, bodyHTML
                        Dim SendResult
                        '******EMAIL VARIABLE******'
                        Set iConf = CreateObject("CDO.Configuration")
                        Set Flds = iConf.Fields
                        With Flds
                            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = "2"
                            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "58.137.229.196"
                            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = "25"
                            .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 600 ' ---- in seconds.
                            .Update
                        End With
                        

                        '******CC email*****'
                        varCCemail="centaraprivilegeclub@chr.co.th" 
                        strCC=varCCemail                            
                        '******CC email*****'

                        Sender = "CentaraSupport <centarasupport@chr.co.th>"
                        Receiver = "niramolim@chr.co.th"'email
                        Set iMsg = CreateObject("CDO.Message")
                        SendResult = False
                        With iMsg
                            Set .Configuration = iConf
                            .BodyPart.Charset = "UTF-8"
                            .To = Receiver
                            .From = Sender
                            .CC= strCC
                            .BCC = "webmaster@chr.co.th, cpc-compbooking@chr.co.th"
                            .BodyPart.Charset = "UTF-8" ' for thai language
                            .Subject = sendMailSubject
        
                            bodyHTML= sendMailToCust
        
                            .CreateMHTMLBody bodyHTML
        
                            '.HTMLBody = bodyHTML
                            Err.Clear
                            .Send
                            If (Err.Number <> 0) Then
                                '***Insert to Log table***'
                                    cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" sending of email to customer rn="&rpID&" failed.', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
                                    Set objExec = Conn.Execute(cmdSql)
                                    Set objExec = Nothing
                                '***Insert to Log table***'
                                Response.Write "Mail sending failed." & Err.Description & "<BR>"

                            Else
                                SendResult = True
                                '***Insert to Log table***'
                                    cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" sending of email to customer rn="&rpID&" completed.', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
                                    Set objExec = Conn.Execute(cmdSql)
                                    Set objExec = Nothing
                                '***Insert to Log table***'
                            End If
                        End With
                        Set iMsg = Nothing

                        If SendResult Then
            
                            Sender = "CentaraSupport <centarasupport@chr.co.th>"
                            Receiver = "niramolim@chr.co.th"'emailSendToResv
                            Set iMsg = CreateObject("CDO.Message")
                            SendResult = False
                            With iMsg
                                Set .Configuration = iConf
                                .BodyPart.Charset = "UTF-8"
                                .To = Receiver
                                .From = Sender
                                .CC= strCC
                                .BCC = "webmaster@chr.co.th, cpc-compbooking@chr.co.th"
                                .BodyPart.Charset = "UTF-8" ' for thai language
                                .Subject = sendMailSubject
            
                                bodyHTML= sendMailToProp
            
                                .CreateMHTMLBody bodyHTML
            
                                '.HTMLBody = bodyHTML
                                Err.Clear
                                .Send
                                If (Err.Number <> 0) Then
                                    '***Insert to Log table***'
                                        cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" sending of email to property rn="&rpID&" failed.', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
                                        Set objExec = Conn.Execute(cmdSql)
                                        Set objExec = Nothing
                                    '***Insert to Log table***'
                                    Response.Write "Mail sending failed." & Err.Description & "<BR>"
                                Else
                                    '***Insert to Log table***'
                                        cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" sending of email to property rn="&rpID&" completed.', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
                                        Set objExec = Conn.Execute(cmdSql)
                                        Set objExec = Nothing
                                    '***Insert to Log table***'
                                    Response.redirect pageSuccess
                                End If
                            End With
                            Set iMsg = Nothing
                        End If
        
                    
                    '************SEND EMAIL*********'
                End If
                
                '******Insert Booking******'

            Else 'Fail
                '***Insert to Log table***'
                    cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" adding of transaction rn="&rpID&" failed.', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
                    Set objExec = Conn.Execute(cmdSql)
                    Set objExec = Nothing
                '***Insert to Log table***'
                Response.redirect "https://www.centarahotelsresorts.com/"&varLangSend&"hmc/pms-fail.asp?guest="&cust_id
            End If  
        '******Inseret to DB******'

    Else
            '***Insert to Log table***'
                cmdSql = "INSERT INTO [dbo].[hmc_tb_r_logs]([log_desc], [active_sts], [created_by], [updated_by]) values('"&Session("hmc_username")&" adding of transaction rn="&rpID&" failed. (no inventory or existing guest_id)', '0', N'"&varIPcreatorx&"', N'"&Session("hmc_username")&"'); "
                Set objExec = Conn.Execute(cmdSql)
                Set objExec = Nothing
            '***Insert to Log table***'
             Response.redirect "https://www.centarahotelsresorts.com/"&varLangSend&"hmc/pms-fail.asp?guest="&cust_id
    End If
    Session("vcid")=""
'Response.write varBookingNo&"<br>"
'Response.write varStatus&"<br>"
'Response.write cmdSql&"<br>"
'Response.write date_arrival&"<br>"
'Response.write date_departure&"<br>"


%>
