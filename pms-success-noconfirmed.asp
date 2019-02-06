<!--Check Login-->
<%'If Session("hmc_login") = True Then%>

<!--#include file ="conf/inc_config.asp"-->
<%
setdropdownlang = True
arrLangSet = array("th","en")
%>
<!DOCTYPE html>
<!--[if IE 7 ]><html class="ie ie7" xml:lang="en-us" lang="en-us> <![endif]-->
<!--[if IE 8 ]><html class="ie ie8" xml:lang="en-us" lang="en-us"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><!--><html class="no-js" xml:lang="en-us" lang="en-us"><!--<![endif]-->
<!-- html header -->

<head>

<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Centara Privilege: Your reservation has been notified.</title>
<meta name="description" content="">
<meta name="subject" CONTENT="">
<meta name="author" content="">
<meta name="copyright" content="">
<meta name="distribution" content="global">
<meta name="robots" content="all, index, follow">
<meta name="rating" content="general">
<link rel="author" href=""/>
<link rel="publisher" href=""/>
<link rel="shortcut icon" href="https://cdn.centarahotelsresorts.com/images/img/favicon.ico">
<!-- hreflang -->

<link rel="alternate" href="https://www.centarahotelsresorts.com/hmc/" hreflang="x-default" />
<link rel="alternate" href="https://www.centarahotelsresorts.com/hmc/" hreflang="en" />

<!-- facebook -->
<meta property="og:title" content="" />
<meta property="og:type" content="article" />
<meta property="og:url" content="https://www.centarahotelsresorts.com/hmc/" />
<meta property="og:image" content="https://www.centarahotelsresorts.com/corporateweb/dist/images/logo-centara-hotels-resorts-200x200.jpg" />
<meta property="og:description" content="" />
<!-- twitter -->
<meta name="twitter:card" content="summary" />
<meta name="twitter:site" content="#explore_chr" />
<meta name="twitter:title" content="">
<meta name="twitter:description" content=""><!--twitter description less than 200 characters-->
<meta name="twitter:creator" content="#explore_chr" />
<meta name="twitter:image" content="https://www.centarahotelsresorts.com/corporateweb/dist/images/logo-centara-hotels-resorts-280x150.jpg" /><!-- size 280x150px -->

<!-- canonical -->
<link rel="canonical" href="https://www.centarahotelsresorts.com/hmc/" />

<!-- Style -->
<link  href="https://www.centarahotelsresorts.com/hmc/assets/bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="https://fonts.googleapis.com/css?family=Open+Sans:300,400,600|Roboto:400,500,700" rel="stylesheet">
<link href="http://fontawesome.io/assets/font-awesome/css/font-awesome.css" rel="stylesheet">
<link href="https://www.centarahotelsresorts.com/hmc/dist/plugin/masterslider/css/masterslider.css" rel="stylesheet" type="text/css">
<link href="https://www.centarahotelsresorts.com/hmc/dist/plugin/masterslider/css/style.css" rel="stylesheet" type="text/css">
<link href="https://www.centarahotelsresorts.com/hmc/dist/plugin/date-picker-v3/css/bootstrap-datepicker.min.css" rel="stylesheet" type="text/css">
<link href="https://www.centarahotelsresorts.com/hmc/dist/style/desktop/main-css.css" rel="stylesheet" type="text/css">
<link href="https://www.centarahotelsresorts.com/hmc/dist/css/pm.css" rel="stylesheet" type="text/css">
<!--HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
<!--[if lt IE 9]>
<script  src="assets/bootstrap/js/html5/html5shiv.min.js"></script>
<script src="assets/bootstrap/js/html5/respond.min.js"></script>
<![endif]-->
<script src="assets/jquery/jquery-1.11.1.min.js"></script>
<script src="assets/bootstrap/js/bootstrap.min.js"></script>


<style>
	.dropdown-menu{margin-top:5px; margin-right:15px;}
	.header-container{z-index:9; position:relative;}
	.info-right{padding-top:20px; padding-bottom:20px; color:#b5121b;}
	.border-box{border: black solid 1px;}
	#banner{margin-top: 30px;}
	.text-form{text-align: left;}
	.contentbox{padding: 10px 0px 0px 0px;}
	.vertical-line{ height: 100%; float: left;  border-left: 1px dotted #444;}
	.bodytxt {padding: 0px 0px;}
	.title-text{font-size: 16px;}
	.text-small{font-size: 11px;}
	
	 @media only screen and (max-width: 480px) {
		div[class=vertical-line] {display:none!important;height:0px;!important;}
	 }
	
	@media only screen and (max-width: 640px) {
		div[class=vertical-line] {display:none!important;height:0px;!important;}
	 }
</style>
</head>
<body style="background-color:#ffffff; color: #000000;" >


<!--Header-->
<header class="header-container" style="z-index:9; position:relative;" >
<div class="container">

	<div class="row">
        <div class="col-sm-12" >

            <div class="col-sm-4">
                <a href="https://www.centarahotelsresorts.com" title="Centara Hotels & Resorts"><img src="https://www.centarahotelsresorts.com/hmc/dist/images/chr-red.png" class="img-responsive"></a>
                <div class="clearfix"></div>
            </div>

            <div class="col-lg-4 col-md-4 col-sm-6 col-xs-12 text-left pull-right" style="padding-top:20px; padding-bottom:20px; color:#b5121b;">
                <div class="contact-right">
    <%

        'Query Information
        rpID = request.querystring("rp")
        strSQL = " SELECT top 1 booking.*, prop.prop_name, prop.prop_name_l_TH, prop.address1 as 'prop_adr', prop.state as 'prop_state', prop.zip as 'prop_zip', prop.phone as 'prop_phone', prop.email_resv as 'prop_email', room.room_name, room.room_name_l_TH, pmprop.contact_email, pmprop.contact_phone from hmc_tb_r_booking booking left join properties prop on(prop.prop_code=booking.prop_code) left join pm_tb_m_properties pmprop on(booking.prop_code=pmprop.prop_code)  left join pm_tb_m_roomtype room on(room.room_code=booking.roomtype  and room.active_sts='1' and booking.prop_code=room.prop_code)  where booking.rn_no='"&rpID&"' and booking.active_sts='1' and booking.sale_sts='P' "

        'Response.write strSQL

        Set objRec = Server.CreateObject("ADODB.Recordset")
        objRec.Open strSQL, Conn, 1,3

            If Not objRec.EOF Then
                intRow = objRec.RecordCount
                If intRow>0 Then

                   'Customer Info.
                    c1c_id = objRec("c1c_id")
                    title = objRec("title")
                    first_name = objRec("first_name")
                    last_name = objRec("last_name")
                    card_name = objRec("card_name")
                    tel = objRec("primary_phone_no")
                    email = objRec("email")
                    address = objRec("address")&" "&objRec("city")&", "&objRec("state")&" "&objRec("country")&" "&objRec("post_code")
                    remark = objRec("note")
                    
                    'Booking Info.
                    conf_id = objRec("conf_id")
                    cvv = objRec("parent_id")
                    creditno = objRec("creditno")
                    prop_code = objRec("prop_code")
                    contact_email = objRec("contact_email")
                    contact_phone = objRec("contact_phone")
                    offer_code = objRec("offer_code")
                    arrival_dt = CDate(objRec("arrival_dt"))
                    departure_dt = CDate(objRec("departure_dt"))
                    prop_name = objRec("prop_name")
                    prop_name_th = objRec("prop_name_l_TH")
                    roomtype = objRec("roomtype")
                    room_name= objRec("room_name")
                    room_name_th= objRec("room_name_l_TH")
                    child = objRec("child")
                    adult = objRec("adult")

                    'Offer Info.
                    card_type="Centara Privilege"
                    offer_price_agreed = objRec("offer_price_agreed")
                    offer_price_show = objRec("offer_price_show")

                    'Email Info.
                    emailSubject="Booking confirmation for "&prop_name&" - "&conf_id

                    'Property
                    prop_adr = objRec("prop_adr")
                    prop_state = objRec("prop_state")
                    prop_zip = objRec("prop_zip")
                    prop_phone = objRec("prop_phone")
                    prop_email = objRec("prop_email")

                    'hmc
                    parent_id=objRec("parent_id")
                    varCCemail="centaraprivilegeclub@chr.co.th" 


                    If session("admin_lang_code") = "TH" Then
                        prop_name = objRec("prop_name_l_TH")
                        lang_code = "th_TH"

                    Else
                        prop_name = objRec("prop_name")
                        lang_code = "en_EN"

                    End If
                 

                End If 'count row
            End If

        objRec.Close()
        Set objRec = Nothing

        'extra request
        If LCase(prop_code)="cgcw" Then
            prop_email="cgcwreservation@chr.co.th"
        ElseIf LCase(prop_choose)="cgc" Then
            emailSendToResv="rsvncgc@chr.co.th"
        End If
    %>
        
                    <strong><%=Session("hmc_username")%> </strong>| <a href="logout.asp"><%=dsC1CForm.item("log_out")%></a><br/>
                    <strong><%=Session("hmc_firstname")%>&nbsp;<%=Session("hmc_lastname")%><br/>
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
        <a href="https://www.centarahotelsresorts.com">
            <img src="https://www.centarahotelsresorts.com/hmc/dist/images/chr-red.png" class="img-responsive" style="margin:0 auto;">
        </a>
    </div>
    <div class="clearfix"></div>

    <div class="col-xs-12 text-center" style="padding-bottom:20px;">
        <strong><%=Session("hmc_username")%> </strong>| <a href="logout.asp"><%=dsC1CForm.item("log_out")%></a><br/>
        <strong><%=Session("hmc_firstname")%>&nbsp;<%=Session("hmc_lastname")%><br/>
    </div>
    <div class="clearfix"></div>
</div>
<!--End Header -->


<!--Card Type Line-->
<section class="headline" style=" text-align: center;">
 	<img src="../hmc/dist/images/voucher/centara-privilege-club-white.png" alt="Centara Privilege" style="height:250px">
</section>

<!--Box Content-->
<section class="bodytxt">
    <div class="container">
        <div class="row">
        <%
        If intRow>0 Then%>
        
        <!--Box Content-->
        <section class="bodytxt">
            <div class="container">
                <div class="row">
                    <p>&nbsp;</p>
                    <p style="font-size: 15px;"><strong><%=dsGeneral.item("your_reservation")%><br />	<%=dsGeneral.item("any_question")%>&nbsp;<%=prop_email%></strong></p>

                    <div class="col-sm-12 contentbox">
                        <div class="col-sm-12" >
                            <div class="col-xs-12 col-sm-6 col-sm-offset-3 border-box text-form title-text" style="padding-top: 10px;padding-bottom: 10px;">
                                <strong><%=dsGeneral.item("hotel_confirmation")%> -</strong>
                            </div>
                            
                            <div class="col-xs-12 col-sm-6 col-sm-offset-3" style="padding-top: 10px;">
                            
                                <div class="col-xs-12 col-sm-3 text-center">	
                                    <p style="font-size:15px;"><strong><%=prop_name%></strong></p>
                                </div>
                                
                                <div class="col-sm-1 " >	
                                    <div class="vertical-line" style="height: 100px;" ></div> 
                                </div>
                                 
                                <div class="col-xs-12 col-sm-8 text-form text-small" style="line-height: 20px;" >	
                                    <div class="row">
                                        <div class="col-xs-5 col-sm-5">
                                         <p><strong><%=dsGeneral.item("name_guest")%><br/>
                                         <%=dsGeneral.item("email_guest")%><br/>
                                         <%=dsGeneral.item("phone_guest")%><br/>
                                         <%=dsGeneral.item("arrival_date")%><br/>
                                         <%=dsGeneral.item("departure_date")%>
                                         </strong></p>
                                        </div>
                                         
                                        <div class="col-xs-7 col-sm-7">
                                            <p><%=title%>&nbsp;<%=first_name%>&nbsp;<%=last_name%><br/>
                                            <%=email%><br/>
                                            <%=tel%><br/>
                                            <%=FormatDateTime(arrival_dt, 1)%><br/>
                                            <%=FormatDateTime(departure_dt, 1)%></p>
                                        </div>
                                        <div class="clearfix"></div>
                                    </div>
                                </div>
                 
                            <div class="clearfix"></div>
                            </div>
                            
                            <div class="col-sm-6 col-sm-offset-3 text-small" style="padding-top: 10px; line-height: 20px; padding-right:0px;">
                            
                                 <div class="col-sm-8 text-form">	
                                  <p>
                                    <strong><%=dsGeneral.item("room_category")%>: </strong><%=room_name%><br/>
                                    <strong><%=dsGeneral.item("no_guest")%>: </strong> <%=dsGeneral.item("adult")%>: <%=adult%>,&nbsp;<%=dsGeneral.item("child")%>: <%=child%><br/>
                                    <strong><%=dsGeneral.item("mode_of_payment")%>:</strong>&nbsp;<%=UCase(card_type)%> (Complimentary Room) <br/>
                                    <strong><%=dsGeneral.item("total_cost")%>:</strong> <%=offer_price_show%> THB <br/>
                                    <strong><%=dsGeneral.item("payment")%>:</strong>	<%=dsGeneral.item("prepaid")%> <br/>
                                    
                                 </p>
                                </div>
                 
                                <div class="col-xs-12 col-sm-4 text-form text-small" style="padding-right:0px;">	
                                    <img src="../hmc/dist/images/voucher/centara-privilege-club-white-500px.png" alt="Centara Privilege" class="img-responsive" style="right: 0;">	 
                                </div>
                                <div class="clearfix"></div>
                                <div class="col-xs-12 text-form text-small">	
                                    <strong><%=dsGeneral.item("special_request")%>:</strong> <%=remark%>	 
                                </div>
                 
                            <div class="clearfix"></div>
                            </div>
                            
                            <div class="col-sm-6 col-sm-offset-3 border-box text-small" style="padding-top: 10px; margin-top: 20px;">
                            
                                 <div class="col-sm-12 text-form" style="">	
                                     <p><strong><%=dsGeneral.item("cancellation_policy")%></strong><br/>
                                                <%=dsGeneral.item("cannot_cancel")%><br/>
                                                <%=dsGeneral.item("no_show")%>
                                    <br/><br/>
                                                <%=dsGeneral.item("physical_voucher")%>
                                    <br/><br/>
                                    <strong>Centara Privilege Club:</strong> Tel. 02-769-1234 Ext 6144-6 | Email centaraprivilegeclub@chr.co.th
                                    <br/><br/>
                                    <strong><%=dsGeneral.item("hotel_reservation")%>:</strong> <br/>
                                    <%=prop_name%><br/>
                                    <%=prop_adr%> <%Response.write prop_state&" "&prop_zip%><br/>
                                    <strong><%=dsGeneral.item("email_guest")%>:</strong> <%=prop_email%><br/>
                                    <strong><%=dsGeneral.item("phone_guest")%>:</strong> <%=prop_phone%> 
                                    </p>
                                </div>
                                 
   
                                  
                            <div class="clearfix"></div>
                            </div>
                 
                            <div class="col-sm-12" style="padding-top: 10px; margin-top: 20px;text-align: center;">
                                 <div class="col-sm-12">	
                                     <input type="button" id="btnBooking" name="btnBooking" value="<%=dsGeneral.item("go_to_ibe_hmc")%>" class="btn btn-sm" style="background-color: #2196F3;border-color:#2196F3;color:#fff;" onclick="window.location.href='https://www.centarahotelsresorts.com/hmc/online-booking/';">
                                    </p>
                                </div>
                            </div>		



                        <div class="clearfix"></div>
                        </div>

                </div>
                <div class="clearfix"></div>
                </div>
            </div>
        </section>

<%
Else
    Response.redirect "https://www.centarahotelsresorts.com/hmc/"
End If
%>
        </div><!--row-->
    </div><!--container-->
</section>

<!--#include file ="inc_file/inc_footer_main.asp"-->



</body>
</html>
<%
'Else
'    url2="index.asp"
'    response.Redirect(url2)
'End If
%>