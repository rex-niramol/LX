<!--Check Login-->
<%If Session("promark_login") = True Then%>

<!--#include file ="conf/inc_config.asp"-->
<%
setdropdownlang = True
arrLangSet = array("th","en")
%>

<head>
<!--#include file ="inc_file/inc_metatag.asp"-->
<!--#include file ="inc_file/inc_style.asp"-->
<!--#include file ="inc_file/inc_script.asp"-->

</head>

<body style="background-color:#ffffff;" >

<%
    guest_id =  Session("guest_number")
%>

<!--dropdown language-->
<!--#include file ="conf/set_condition_language.asp"-->
<!--#include file ="inc_file/inc_dropdownlang.asp"-->

<!--Header-->
<header class="header-container" style="z-index:9; position:relative;" >
<div class="container">

	<div class="row">
    <div class="col-sm-12" >
		<div class="col-sm-4">
    	<a href="http://www.centarahotelsresorts.com/" title="Centara Hotels & Resorts"><img src="https://www.centarahotelsresorts.com/ultimate-dining/dist/images/chr-red.png" class="img-responsive"></a>
        <div class="clearfix"></div>
        </div>
        <div class="col-lg-4 col-md-4 col-sm-6 col-xs-12 text-left pull-right" style="padding-top:20px; padding-bottom:20px; color:#b5121b;">
			<div class="contact-right">
<%

'Query Information
rpID = request.querystring("rp")
strSQL = " SELECT top 1 booking.*, prop.prop_name, prop.prop_name_l_TH, prop.address1 as 'prop_adr', prop.state as 'prop_state', prop.zip as 'prop_zip', prop.phone as 'prop_phone', prop.email_resv as 'prop_email', room.room_name, room.room_name_l_TH, offer.offer_price_agreed, offer.offer_price_show, offer.offer_name, offer.offer_name_l_TH, offer.offer_ads, offer.offer_ads_l_TH, offer.offer_cond, offer.offer_cond_l_TH, offer.offer_photo_thumb, offer.offer_photo_full, offer.note as 'offer_note', pmprop.contact_email, pmprop.contact_phone from pm_tb_r_booking booking left join pm_tb_r_sales sales on(sales.guest_id=booking.guest_id)  left join properties prop on(prop.prop_code=booking.prop_code) left join pm_tb_m_properties pmprop on(booking.prop_code=pmprop.prop_code)  left join pm_tb_m_roomtype room on(room.room_code=booking.roomtype and room.active_sts='1' and booking.prop_code=room.prop_code)  left join pm_tb_a_offers offer on(offer.prop_code=sales.prop_code and offer.offer_code=booking.offer_code) where booking.rn_no='"&rpID&"' and booking.guest_id='"&Session("guest_number")&"' and booking.active_sts='1' and booking.sale_sts='C' "

Set objRec = Server.CreateObject("ADODB.Recordset")
objRec.Open strSQL, Conn, 1,3

    If Not objRec.EOF Then
        intRow = objRec.RecordCount
        If intRow>0 Then
            title = objRec("title")
            first_name = objRec("first_name")
            last_name = objRec("last_name")
            email = objRec("email")
            mobile_no = objRec("mobile_no")
            card_type = objRec("card_type")

           'Customer Info.
            c1c_id = objRec("c1c_id")
            passport_id = objRec("passport_id")
            cer_id = objRec("cer_id")
            gender = objRec("gender")
            nationality = objRec("nationality")
            title = objRec("title")
            first_name = objRec("first_name")
            last_name = objRec("last_name")
            card_name = objRec("card_name")
            tel = objRec("mobile_no")
            email = objRec("email")
            address = objRec("address1")&" "&objRec("city")&", "&objRec("state")&" "&objRec("country")&" "&objRec("post_code")
            remark = objRec("note")
            
            'Booking Info.
            conf_id = objRec("conf_id")
            cvv = objRec("parent_id")
            creditno = objRec("creditno")
            prop_code = objRec("prop_code")
            contact_email = objRec("contact_email")
            contact_phone = objRec("contact_phone")
            offer_code = objRec("offer_code")
            card_type = objRec("card_type")
            arrival_dt = objRec("arrival_dt")
            arrArrival=Split(arrival_dt,"-")
            departure_dt = objRec("departure_dt")
            arrDeparture=Split(departure_dt,"-")
            prop_name = objRec("prop_name")
            prop_name_th = objRec("prop_name_l_TH")
            roomtype = objRec("roomtype")
            room_name= objRec("room_name")
            room_name_th= objRec("room_name_l_TH")
            child = objRec("child")
            adult = objRec("adult")

            'Offer Info.
            offer_price_agreed = objRec("offer_price_agreed")
            offer_price_show = objRec("offer_price_show")
            offer_name = objRec("offer_name")
            offer_name_l_TH = objRec("offer_name_l_TH")
            offer_ads = objRec("offer_ads")
            offer_ads_l_TH = objRec("offer_ads_l_TH")
            offer_cond = objRec("offer_cond")
            offer_cond_l_TH = objRec("offer_cond_l_TH")
            offer_img_front = objRec("offer_photo_thumb")
            offer_img_back = objRec("offer_photo_full")
            offer_note = objRec("offer_note")

            'Email Info.
            emailSubject="Booking confirmation for "&prop_name&" - "&conf_id

            'Property
            prop_adr = objRec("prop_adr")
            prop_state = objRec("prop_state")
            prop_zip = objRec("prop_zip")
            prop_phone = objRec("prop_phone")
            prop_email = objRec("prop_email")

            If session("admin_lang_code") = "TH" Then
                prop_name = objRec("prop_name_l_TH")
                lang_code = "th_TH"
            Else
                prop_name = objRec("prop_name")
                lang_code = "en_EN"
            End If
         
            'Check Card Type
            If card_type = "gold" then
                card_type = dsGeneral.item("gold_card")
                card_title = "https://www.centarahotelsresorts.com/ultimate-dining/dist/images/gold-card.png"

            Else 
                card_type = dsGeneral.item("platinum_card")
                card_title = "https://www.centarahotelsresorts.com/ultimate-dining/dist/images/platinum-card.jpg"
            End If

        End If 'count row
    End If
    objRec.Close()
    Set objRec = Nothing
%>
    
<strong><%=first_name%> </strong>| <a href="https://www.centarahotelsresorts.com/ultimate-dining/logout.asp"><%=dsC1CForm.item("log_out")%></a><br/>
<strong><%=dsC1CForm.item("id")%> : <%=guest_id%><br/>
<span style="color:#857650;"><%=card_type%></span></strong><br/>
<strong><span style="color:#000000;"><%=prop_name%></span></strong>
			</div>
            <div class="clearfix"></div>
			</div>
                 

    <div class="clearfix"></div>

</div>

<div class="clearfix"></div>
</div>

</div>

<div class="clearfix"></div>
</div>
</header>


<!--For Mobile -->
<div class="header-sm"> 

<div class="col-xs-6 col-sm-12 text-center">
<a href="http://www.centarahotelsresorts.com/">
<img src="https://www.centarahotelsresorts.com/ultimate-dining/dist/images/chr-red.png" class="img-responsive" style="margin:0 auto;">
</a>
</div>

<div class="clearfix"></div>

<div class="col-xs-12 text-center" style="padding-bottom:20px;">
<strong><%=first_name%> </strong>| <a href="https://www.centarahotelsresorts.com/th/ultimate-dining/logout.asp">LOG OUT</a><br/>
<strong>ID : <%=guest_id%><br/>
<span style="color:#857650;"><%=card_type%></span></strong><br/>
<strong><span style="color:#000000;"><%=prop_name%></span></strong>

</div>

<div class="clearfix"></div>

</div>
<!--End Header -->


<!--Card Type Line-->
<section class="headline" >
 		<img src="<%=card_title%>" alt="Gold Card"/ class="img-responsive" style="margin:0 auto;">
<div class="clearfix"></div>
</section>

<!--Box Content-->
<section class="bodytxt">
<div class="container">
    <div class="row">
    <%
    If intRow>0 Then%>
    
        <table width='100%' border='0' cellpadding='0' cellspacing='0' align='center'>
            <tr>
                <td width='80%' valign='top' style='padding:20px'>
                    <!-- Start Header-->
                    <table width='80%' border='0' cellpadding='0' cellspacing='0' align='center' class='deviceWidth' bgcolor='#ffffff'>
                        <tr>
                            <td colspan="2" width='100%' style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                           <p  align='center'><strong>Thank you, your booking is now confirmed
                           <br>A confirmation email has been sent to <%=email%></strong></p>
                           <p>&nbsp;</p>
                           </td>

                        <tr>
                            <td width='40%' style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Hotel confirmation
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                <%=conf_id%>
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Name of the guest
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                               <%=title%>&nbsp;<%=first_name%>&nbsp;<%=last_name%>
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Email of the guest
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                               <%=email%>
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Phone of the guest
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                               <%=tel%>
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                No. of the guest(s)
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                               <%=adult+child%>&nbsp;Pax
                           </td>
                       </tr>
                        <tr>
                            <td  style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Number of room(s)
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                1&nbsp;<%=room_name%>
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Arrival date
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                <%=FormatDateTime(arrival_dt, 1)%>
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Departure date
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                <%=FormatDateTime(departure_dt, 1)%>
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Property
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                <%=prop_name%>
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Room category
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                <%=room_name%>
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Mode of payment
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Promark Voucher (<%=offer_name%>)
                           </td>
                       </tr>
                        <tr>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Total Cost
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                <%=offer_price_show%> THB
                           </td>
                       </tr>
                        <tr>
                            <td  style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                Payment
                           </td>
                            <td style='padding-top:0px; padding-bottom:0px; padding-left:35px; padding-right:35px; color:#000000;'>
                                check-in counter
                           </td>
                       </tr>
                        <tr valign='top' bgcolor='#ecebeb'>
                            <td colspan="2" width='100%' style='padding-top:20px; padding-bottom:20px; padding-left:35px; padding-right:35px; color:#000000;'>
                            <p><strong>Cancellation Policy</strong><br>
                            This booking cannot be cancelled. In order to amend this booking please contact the hotel directly.<br>In case of no-show, a penalty of full stay will be applied.</p>
                            <p>Please bring your physical voucher along with your member card to check-in. Without physical voucher, a penalty of full stay will be applied.</p>
                            <p><strong>Promark Hotline:  <%=offer_note%></strong></p>
                            <p><strong>Hotel Reservation Enquiries:</strong>
                            <br><%=prop_name%>
                            <br><%=prop_adr%>
                            <br><%=prop_state&" "&prop_zip%>
                            <br>Email: <%=prop_email%>
                            <br>Phone: <%=prop_phone%>
                        </tr>

                        <tr valign='top' bgcolor='#ecebeb'>
                            <td colspan="2" width='100%' style='padding-top:20px; padding-bottom:20px; padding-left:35px; padding-right:35px; color:#000000;'>
                                <img src="<%=offer_img_front%>" alt="<%=offer_name%>">
                            </td>
                        </tr>

                        <tr>
                            <td colspan="2" style="text-align:right;">
                                <input type="button" id="btnBooking" name="btnBooking" value="Go to my vouchers" class="btn btn-sm" style="background-color: #2196F3;border-color:#2196F3;color:#fff;" onclick="window.location.href='https://www.centarahotelsresorts.com/ultimate-dining/voucher.asp';">
                           </td>
                       </tr>
                    </table><!-- End Header -->



                    <!-- Start body-->
                    <table width='600' border='0' cellpadding='0' cellspacing='0' align='center' class='deviceWidth' bgcolor='#ffffff'>

                    </table><!-- End body -->
              </td>
            </tr>
        </table> <!-- End Wrapper -->

    <%
    Else
        Response.redirect "https://www.centarahotelsresorts.com/ultimate-dining/voucher.asp"
    End If
    %>
    </div><!--row-->
</div>
</section>

<!--#include file ="inc_file/inc_footer_main.asp"-->
<!--#include file ="inc_file/inc_remarketing.asp"-->


</body>
</html>
<%
Else
    url2="index.asp"
    response.Redirect(url2)
End If
%>