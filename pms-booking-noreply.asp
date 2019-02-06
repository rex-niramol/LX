<!--#include file ="conf/inc_config.asp"-->
<%
	If session("admin_lang_code") = "TH" Then
        varLangSend="th/"
	Else
        varLangSend=""
	End If

    offer_code=request.Form("offercode")
    card_code=request.Form("card_code")
    prop_code = request.Form("prop_code")
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
    varAllot=2
    pmsSystem="COMANCHE"
    pmsSystemName="COMANCHE"
    remark=request.Form("remark")&" (PROMARK, MEM: "&Session("guest_number")&" EXP: "&Session("date_check")&")"

    varRecPriceAgreed="0"
    varRecPriceShow="0"
    varBookingSts="C"

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

    If varAllot>0 Then
        '******Pricing******'
            cmdSql = " SELECT offer_price_show, offer_price_agreed, hotels_availability from pm_tb_a_offers_pricing where active_sts='1' and offer_code='"&offer_code&"' "

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
            If prop_code=prop_choose Then
                market_segment="COMP"
                varRecPriceAgreed="0"
                varRecPriceShow="0"
            Else
                market_segment="LOYAL"
            End If
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
            End If
        '******Property Info******'

        '******PMSsystem******'
            cmdSql = " SELECT sys_code, sys_name, hotels_avai from pm_tb_m_pms where active_sts='1' "

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

        '******CommandChe Web Service
            Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
            oXmlHTTP.Open "POST", "http://www.chrnetwork.com/promark/ProMark.asmx", False 

            oXmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8" 
            oXmlHTTP.setRequestHeader "SOAPAction", "http://tempuri.org/saveBooking"

            SOAPRequest ="<?xml version=""1.0"" encoding=""utf-8""?>" & _
              "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
              "<soap:Body>" & _
                "<saveBooking xmlns=""http://tempuri.org/"">" &_
                      "<property>"&prop_choose&"</property>" &_
                      "<title>"&title&"</title>" &_
                      "<first_name>"&firstname&"</first_name>" &_
                      "<last_name>"&lastname&"</last_name>" &_
                      "<email>"&email&"</email>" &_
                      "<phone>"&phone&"</phone>" &_
                      "<arrive>"&date_arrival&"</arrive>" &_
                      "<depart>"&date_departure&"</depart>" &_
                      "<room_night>1</room_night>" &_
                      "<room_type>"&roomtype&"</room_type>" &_
                      "<adult>"&adult&"</adult>" &_
                      "<child>"&child&"</child>" &_
                      "<credit_no>"&creditPms&"</credit_no>" &_
                      "<price>"&varRecPriceAgreed&"</price>" &_
                      "<business_code>"&offer_code&"</business_code>" &_
                      "<guest_id>"&Session("guest_number")&"</guest_id>" &_
                      "<pms>"&pmsSystem&"</pms>" &_
                      "<card_exp>"&expCreditPMS&"</card_exp>" &_
                      "<card_type>"&card_code&"</card_type>" &_
                      "<market>"&market_segment&"</market>" &_
                      "<remark>"&remark&"</remark>" &_
                      "<card_owner>"&UCase(prop_code)&"</card_owner>" &_
                "</saveBooking>" &_
              "</soap:Body>" &_
              "</soap:Envelope>"

            oXmlHTTP.send SOAPRequest  
            varResponse= oXmlHTTP.responseText
            function findStringBetween(s_string, s_pre, s_post, n_options)
              ' n_options:
              '  1: IGNORE case in the search
              dim a, b
              ' init
              findStringBetween = ""
              ' find pre-word
              a = instr(1, s_string, s_pre, n_options)
              if a>0 then
                a = a + len(s_pre)
                b = instr(a, s_string, s_post, n_options)
                if b>0 and not (s_post = "") then
                  findStringBetween = mid(s_string, a, b-a)
                else
                  findStringBetween = mid(s_string, a)
                end if
              end If
            end function  

            varBookingNo = findStringBetween(varResponse, "<reservno>", "</reservno>", 0)
            varStatus = findStringBetween(varResponse, "<status>", "</status>", 0)
        '******CommandChe Web Service

        '******Inseret to DB******'
            varSuccess=InStr(LCase(varResponse),"success")
            If LCase(varStatus)="success" Then 'Success
                If UCase(pmsSystemName)<>"COMANCHE" Then
                     varBookingSts="P"
                End If

                '******Insert Booking******'
                cmdSql = "INSERT INTO [dbo].[pm_tb_r_booking]([guest_id], [c1c_id], [passport_id], [cer_id], [conf_id], [prop_code], [prop_system], [offer_code], [offer_price_agreed], [offer_price_show], [card_type], [arrival_dt], [departure_dt], [roomtype],	[adult], [child], [creditno], [creditexp], [lang_ref], [email], [card_name], [title], [first_name], [last_name], [gender], [nationality], [marital], [birthday], [mobile_no], [home_no], [office_no], [primary_phone_no], [fax], [company], [address1], [address2], [city], [state], [country], [post_code], [note], [sale_sts], [active_sts], [delete_sts], [value_seq], [parent_id], [created_by], [updated_by]) (select top 1 [guest_id], [c1c_id], [passport_id], [cer_id], '"&varBookingNo&"', '"&prop_choose&"', '"&pmsSystemName&"', '"&offer_code&"', '"&varRecPriceAgreed&"', '"&varRecPriceShow&"', '"&card_code&"', '"&varArrDate&"', '"&varDptDate&"', '"&roomtype&"', '"&adult&"', '"&child&"', '"&creditStore&"', '"&expCredit&"', [lang_ref], '"&email&"', [card_name], '"&title&"', '"&firstname&"', '"&lastname&"', [gender], [nationality], [marital], [birthday], [mobile_no], [home_no], [office_no], '"&phone&"', [fax], [company], [address1], [address2], [city], [state], [country], [post_code], '"&remark&"', '"&varBookingSts&"', '1', 'N', '99', '"&prop_code&"', '"&Session("guest_number")&"', '"&Session("guest_number")&"' from pm_tb_r_sales where guest_id ='"&Session("guest_number")&"'); "
                stringsql=cmdSql
                rpID = "1"

                'cmdSql = cmdSql & "Set Nocount no;"
                'cmdSql = cmdSql & "SELECT NewsID = SCOPE_IDENTITY();"
                'cmdSql = cmdSql & "SELECT SCOPE_IDENTITY() AS NewsID;"
                'cmdSql = cmdSql & "SELECT @@IDENTITY AS 'NewsID';"
                'cmdSql = cmdSql & "Set Nocount off;"

                Set objExec = Conn.Execute(cmdSql)
                Set objExec = Nothing
                If Err.Number = 0 Then
                    cmdSql = "SELECT MAX(rn_no) AS NewsID from pm_tb_r_booking;"
                    Set recNewsID = Server.CreateObject("ADODB.Recordset")
                    recNewsID.Open cmdSql, Conn, 1,3
                    If Not recNewsID.EOF then
                        rpID = recNewsID("NewsID")
                    End If
                    recNewsID.Close()
                    Set recNewsID = Nothing
                    'cmdSQL="update pm_tb_r_booking set conf_id='"&varBookingNo&"' where rn_no='"&rpID&"' "
                    'conn.Execute cmdSQL

                    '************SEND EMAIL*********'
                        If UCase(pmsSystem)="COMANCHE" Then
                            pageSuccess="http://www.centarahotelsresorts.com/"&varLangSend&"promark/pms-success-confirmed.asp?rp="&rpID
                            sendMailToCust="http://www.centarahotelsresorts.com/"&varLangSend&"promark/inc_file/sendemail-customer-pms-mhtml-comp-confirmed.asp?rp="&rpID
                            sendMailToProp="http://www.centarahotelsresorts.com/"&varLangSend&"promark/inc_file/sendemail-property-pms-mhtml-comp-confirmed.asp?rp="&rpID
                            sendMailSubject=varBookingNo & " - Booking confirmation for " & emailPropName
                        Else
                             pageSuccess="http://www.centarahotelsresorts.com/"&varLangSend&"promark/pms-success-noconfirmed.asp?rp="&rpID
                            sendMailToCust="http://www.centarahotelsresorts.com/"&varLangSend&"promark/inc_file/sendemail-customer-pms-mhtml-comp-noconfirmed.asp?rp="&rpID
                            sendMailToProp="http://www.centarahotelsresorts.com/"&varLangSend&"promark/inc_file/sendemail-property-pms-mhtml-comp-noconfirmed.asp?rp="&rpID
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
                            Select Case prop_code
                                Case "cglb"
                                    varCCemail="goldcard@chr.co.th"
                                Case "chy"
                                    varCCemail="goldcardchy@chr.co.th"
                                Case "cgcw"
                                    varCCemail="platinumcard@chr.co.th"
                                Case "cmbr"
                                    varCCemail="platinumcardcmbr@chr.co.th"
                                Case "cud"
                                    varCCemail="goldcardcud@chr.co.th"
                                Case Else
                                    varCCemail="memberservices@centara1card.com" 
                            End Select
                            'strCC="ArthidCh@chr.co.th, sorayase@chr.co.th, verasakon@chr.co.th" memberservices@centara1card.com
                            strCC=varCCemail                            
                        '******CC email*****'

                        Sender = "noreply <noreply@chr.co.th>"
                        Receiver = email
                        Set iMsg = CreateObject("CDO.Message")
                        SendResult = False
                        With iMsg
                            Set .Configuration = iConf
                            .BodyPart.Charset = "UTF-8"
                            .To = Receiver
                            .From = Sender
                            .CC= strCC
                            .BCC = "webmaster <webmaster@chr.co.th>"
                            .BodyPart.Charset = "UTF-8" ' for thai language
                            .Subject = sendMailSubject
        
                            bodyHTML= sendMailToCust
        
                            .CreateMHTMLBody bodyHTML
        
                            '.HTMLBody = bodyHTML
                            Err.Clear
                            .Send
                            If (Err.Number <> 0) Then
                                Response.Write "Mail sending failed." & Err.Description & "<BR>"
                            Else
                                SendResult = True
                            End If
                        End With
                        Set iMsg = Nothing

                        If SendResult Then
            
                            Sender = "noreply <noreply@chr.co.th>"
                            Receiver = emailSendToResv
                            Set iMsg = CreateObject("CDO.Message")
                            SendResult = False
                            With iMsg
                                Set .Configuration = iConf
                                .BodyPart.Charset = "UTF-8"
                                .To = Receiver
                                .From = Sender
                                '.CC= strCC
                                .BCC = "webmaster <webmaster@chr.co.th>"
                                .BodyPart.Charset = "UTF-8" ' for thai language
                                .Subject = sendMailSubject
            
                                bodyHTML= sendMailToProp
            
                                .CreateMHTMLBody bodyHTML
            
                                '.HTMLBody = bodyHTML
                                Err.Clear
                                .Send
                                If (Err.Number <> 0) Then
                                    Response.Write "Mail sending failed." & Err.Description & "<BR>"
                                Else
                                    Response.redirect pageSuccess
                                End If
                            End With
                            Set iMsg = Nothing
                        End If
        
                    
                    '************SEND EMAIL*********'
                End If
                
                '******Insert Booking******'

            Else 'Fail
                Response.redirect "http://www.centarahotelsresorts.com/"&varLangSend&"promark/pms-fail.asp"
            End If  
        '******Inseret to DB******'

    Else
             Response.redirect "http://www.centarahotelsresorts.com/"&varLangSend&"promark/pms-fail.asp"
    End If
    Session("vcid")=""
'Response.write varBookingNo&"<br>"
'Response.write varStatus&"<br>"
'Response.write cmdSql&"<br>"
'Response.write date_arrival&"<br>"
'Response.write date_departure&"<br>"

%>
