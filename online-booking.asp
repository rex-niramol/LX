<!--Check Login-->
<%If Session("hmc_login") = True Then%>

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
<title>Centara Privilege</title>
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

<link rel="alternate" href="../hmc/" hreflang="x-default" />
<link rel="alternate" href="../hmc/" hreflang="en" />

<!-- facebook -->
<meta property="og:title" content="" />
<meta property="og:type" content="article" />
<meta property="og:url" content="https://www.centarahotelsresorts.com/hmc/" />
<meta property="og:image" content="https://www.centarahotelsresorts.com/corporateweb/dist/images/logo-centara-hotels-resorts-200x200.jpg" />
<meta property="og:description" content="" />
<!-- twitter -->
<meta name="twitter:card" content="summary" />
<meta name="twitter:site" content="#explorecentara" />
<meta name="twitter:title" content="">
<meta name="twitter:description" content=""><!--twitter description less than 200 characters-->
<meta name="twitter:creator" content="#explorecentara" />
<meta name="twitter:image" content="https://www.centarahotelsresorts.com/corporateweb/dist/images/logo-centara-hotels-resorts-280x150.jpg" /><!-- size 280x150px -->

<!-- canonical -->
<link rel="canonical" href="../hmc/" />

<!-- Style -->
<link  href="https://www.centarahotelsresorts.com/hmc/assets/bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="https://fonts.googleapis.com/css?family=Open+Sans:300,400,600|Roboto:400,500,700" rel="stylesheet">

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
<script src="https://www.centarahotelsresorts.com/hmc/assets/jquery/jquery-1.11.1.min.js"></script>
<script src="https://www.centarahotelsresorts.com/hmc/assets/bootstrap/js/bootstrap.min.js"></script>


<style>
	.nav-style{z-index:999; padding-right:15px; padding-top:15px; right:0; position:absolute;}
	.dropdown-menu{margin-top:5px; margin-right:15px;}
	.header-container{z-index:9; position:relative;}
	.info-right{padding-top:20px; padding-bottom:20px; color:#b5121b;}
	.border-box{border: black solid 1px;}
	#banner{margin-top: 30px;}
	.text-form{text-align: left;}
	.contentbox{padding: 40px 0px 0px 0px;}
</style>
</head>

<body style="background-color:#ffffff;" >
<!--dropdown language-->
<!--#include file ="conf/set_condition_language.asp"-->
<!--#include file ="inc_file/inc_dropdownlang.asp"-->
<%

	If session("admin_lang_code") = "TH" Then
        varLangSend="th/"
	Else
        varLangSend=""
	End If

    offer_hotels= "'cgcw','cglb','chbr','ckbr','cmbr','cpbr','csbr','cgpx','ckr','ckt','cvp','csv','cmp','cak','cwb','csk','ckc','crr','kpc','wsp','cms','chy','nvp','cpy','cct','cud','cgc','csb','cah','cmj','cbs','cap','czsc'"

%>

<!--dropdown language-->


    <!--Header-->
    <header class="header-container">
        <div class="container">

            <div class="row">
                <div class="col-sm-12" >
                    <div class="col-sm-4">
                        <a href="<%=chr_link%>" title="Centara Hotels & Resorts"><img src="../dist/images/chr-red.png" class="img-responsive"></a>
                        <div class="clearfix"></div>
                    </div>
                    <div class="col-lg-4 col-md-4 col-sm-6 col-xs-12 text-left pull-right" style="padding-top:20px; padding-bottom:20px; color:#b5121b;">
                        <div class="contact-right">
                            <strong><%=Session("hmc_username")%> </strong>| <a href="../logout.asp"><%=dsC1CForm.item("log_out")%></a><br/>
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
    <!--End Header -->


    <!--Box Content-->
    <section class="bodytxt">
        <div class="container">
            <div class="row">

               
                <h3 style="color:#1c1b19;"><strong><%=dsGeneral.item("check_room_avai")%></strong></h3> 
                
                <form id="frmOnlineBooking"  name="frmOnlineBooking" action="../pms-booking/" method="post" onsubmit="return validateForm()">
                    <!--Room Booking Information-->
                    <div class="col-sm-12 contentbox">
                        <section id="filter">
                            <%
                                todayDate=Date
                                tomorrow=DateAdd("d", 1, todayDate)
                                varToDay=Right("0"&Day(todayDate),2)&"/"&Right("0"&Month(todayDate),2)&"/"&Year(todayDate)
                                varTomorrow=Right("0"&Day(tomorrow),2)&"/"&Right("0"&Month(tomorrow),2)&"/"&Year(tomorrow)
                            %>
                            <div class="col-sm-12" >
                                <div class="col-sm-10 col-sm-offset-1 border-box" style="padding:20px;">

                                    <div class="col-sm-6">
                                           <div class="form-group text-form">
                                               <label for="hotels" >*<%=dsGeneral.item("centara_hotels_and_resorts")%></label>
                                               <select name="property" id="property" class="form-control" onchange="javascript:showRoomtype(this.form);">
                                                <option value='' >--- <%=dsGeneral.item("please_select")%> ---</option>
                                                <%
                                                '******Properties******' 
                                                cmdProp = " SELECT prop_code, prop_name from properties where activate='1' and prop_code in("&offer_hotels&") order by prop_name"

                                                Set recProp = Server.CreateObject("ADODB.Recordset")
                                                recProp.Open cmdProp, Conn, 1,3

                                                While Not recProp.EOF
                                                    Response.write "<option value='" & recProp("prop_code") & "' >" & recProp("prop_name") & "</option>"
                                                recProp.MoveNext
                                                Wend
                                                recProp.Close()
                                                Set recProp = Nothing
                                                '******Properties******'

                                                %>
                                                </select>
                                           </div>
                                    </div>
                                      
                                    <div class="col-sm-6">
                                          <div class="form-group text-form">
                                            <label for="room_type" >*<%=dsGeneral.item("room_type")%></label>
                                                <div id="show-roomtype">
                                                    <select name='roomtype' id='roomtype' class='form-control'>
                                                        <option value='' >--- <%=dsGeneral.item("please_select")%> ---</option>
                                                        <!--**************SHOW ROOMTYPE**************-->

                                                        <!--**************SHOW ROOMTYPE**************-->
                                                    </select>
                                                </div>
                                           </div>
                                    </div>
                                    <div class="clearfix"></div>

                                    <div class="col-sm-3">
                                        <div class="form-group text-form">
                                            <label for="arrival_date" >*<%=dsGeneral.item("arrival_date")%> (dd/mm/yyyy)</label>
                                           <input type="text"  placeholder="Arrival Date" name="arrival" id="arrival" value="<%=varTomorrow%>" style="background:url(<%=domain%>dist/images/icon/calendar.png) right center no-repeat #fff; border-right:none;width:100%" class="form-control" onchange="javascript:showDeparture(this.form);">
                                        </div>
                                    </div>
                                          
                                    <div class="col-sm-3">
                                        <div class="form-group text-form">
                                            <label for="dep_date" >*<%=dsGeneral.item("departure_date")%> (<%=dsGeneral.item("one_night_only")%>)</label>
                                            <input type="text" id="departure" name="departure" class="form-control" readonly>
                                        </div>
                                    </div>
                                          
                                    <div class="col-sm-3">
                                        <div class="form-group text-form">
                                            <label for="adulttxt" >*<%=dsGeneral.item("adult")%></label>
                      
                                            <select id="adult" name="adult" class="form-control" >
                                                <option value="1">1</option>
                                                <option value="2" selected>2</option>
                                            </select>
                                        </div>
                                    </div>
                                          
                                    <div class="col-sm-3">
                                        <div class="form-group text-form">
                                            <label for="childtxt" >*<%=dsGeneral.item("child")%></label>
                                            <select name="child" id="child" class="form-control" >
                                                <option value="0">0</option>
                                                <option value="1">1</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div class="clearfix"></div>
                                         
                                    <div class="col-sm-9" id="sale-form-matrix">

                                        <!--**************INCLUDE CONTENT**************-->

                                        <!--**************INCLUDE CONTENT**************-->                            

                                    </div>
                                          
                                    <div class="col-sm-3" style="text-align: right; padding-bottom: 20px;">
                                        <input type="button" id="btnChkAvi" name="btnChkAvi" value="<%=dsGeneral.item("check_availibility")%>" class="btn btn-blue"  onclick="javascript:funChkAvi(this.form);" >
                                    </div> 
                                    <div class="clearfix"></div>

                                </div>
                            </div><!--End Filter-->

                        </section>
                    </div>
                    <!--Room Booking Information-->

                    <!--Guest Information-->
                        <div class="col-sm-12" style="margin-top: 20px;" >

                        <h3 style="color:#1c1b19;"><strong><%=dsGeneral.item("guest_information")%></strong></h3>
                            <section id="confirm">
                                <div class="col-sm-10 col-sm-offset-1" style="border: black solid 1px;padding:20px;">

                                    <div class="col-sm-6 col-md-3">
                                        <div class="form-group text-form">
                                           <label for="membership_no" >*<%=dsC1CForm.item("member_id")%></label>  
                                            <input type="text" name="membership_no" id="membership_no" class="form-control">
                                        </div>
                                    </div>
                                    <div class="col-sm-6 col-md-3">
                                        <div class="form-group text-form">
                                            <label for="expiry_date" >*<%=dsGeneral.item("expiry_date")%> (dd/mm/yyyy)</label>
                                           <input type="text"  placeholder="Expiry Date" name="expiry_date" id="expiry_date" value="<%=varTomorrow%>" style="background:url(<%=domain%>dist/images/icon/calendar.png) right center no-repeat #fff; border-right:none;width:100%" class="form-control">
                                        </div>
                                    </div>

                                    <div class="col-sm-2 col-md-2">
                                        <div class="form-group text-form">
                                            <label for="redemption_type" >*<%=dsGeneral.item("redeem_type")%></label>
                      
                                            <select id="redeem_type" name="redeem_type" class="form-control" >
                                                <option value="HMC App.">APP.</option>
                                                <option value="Call Center">Call Center</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div class="col-sm-4 col-md-2">
                                        <div class="form-group text-form">
                                           <label for="mobile_code" ><%=dsGeneral.item("redeem_code")%></label> 
                                            <input type="text" name="mobile_code" id="mobile_code" class="form-control">
                                        </div>
                                    </div>
                                    <div class="col-sm-4 col-md-2">
                                        <div class="form-group text-form">
                                           <label for="hmc_code" ><%=dsGeneral.item("response_code")%></label>
                                            <input type="text" name="hmc_code" id="hmc_code" class="form-control">
                                        </div>
                                    </div>
                                    <div class="clearfix"></div>
                                          


                                
                                    <div class="col-sm-2">
                                       <div class="form-group text-form">
                                           <label for="pre_name" >&nbsp;</label>  
                                           <select name="title" id="title" class="form-control" >
                                                <option value='KHUN'>Khun</option>
                                                <option value='MR.'>Mr.</option>
                                                <option value='MRS.'>Mrs.</option>
                                                <option value='MS.'>Ms.</option>       
                                           </select>
                                        </div>
                                    </div>
                                  
                                    <div class="col-sm-4">
                                        <div class="form-group text-form">
                                           <label for="firstname" >*<%=dsGeneral.item("first_name")%></label>  
                                            <input type="text" name="firstname" id="firstname" class="form-control">
                                        </div>
                                    </div>
                                      
                                    <div class="col-sm-6">
                                        <div class="form-group text-form">
                                          <label for="lastname" >*<%=dsGeneral.item("last_name")%> &nbsp;<%=dsGeneral.item("coupon_reserved_members_only")%></label>
                                           <input type="text" name="lastname" id="lastname" class="form-control">
                                        </div>
                                    </div>
                                    <div class="clearfix"></div>
                                         
                                    <div class="col-sm-6">
                                        <div class="form-group text-form">
                                            <label for="email" >*<%=dsGeneral.item("email_guest")%></label>
                                            <input type="email" name="email" id="email" class="form-control">
                                        </div>
                                    </div>
                                          
                                    <div class="col-sm-6">
                                        <div class="form-group text-form">	
                                            <label for="phone" ><%=dsGeneral.item("phone_guest")%></label>
                                            <input type="text" id="phone" name="phone" class="form-control" value="">
                                        </div>
                                    </div>
                                         
                                    <div class="col-sm-6">
                                        <div class="form-group text-form">
                                            <div class="row">
                                                <div class="col-sm-3">
                                                    <label for="credit_no" >*<%=dsGeneral.item("credit_no")%></label>
                                                    <input type="text" id="crd01" name="crd01" class="form-control" maxlength="4" onKeyup="autotab(this, document.frmOnlineBooking.crd02)" onkeypress="return event.charCode >= 48 && event.charCode <= 57">
                                                </div>
                                                <div class="col-sm-3">
                                                    <label for="credit_no" >&nbsp;</label>
                                                    <input type="text" id="crd02" name="crd02" class="form-control" maxlength="4" onKeyup="autotab(this, document.frmOnlineBooking.crd03)" onkeypress="return event.charCode >= 48 && event.charCode <= 57">
                                                </div>
                                                <div class="col-sm-3">
                                                    <label for="credit_no" >&nbsp;</label>
                                                    <input type="text" id="crd03" name="crd03" class="form-control" maxlength="4" onKeyup="autotab(this, document.frmOnlineBooking.crd04)" onkeypress="return event.charCode >= 48 && event.charCode <= 57">
                                                </div>
                                                <div class="col-sm-3">
                                                    <label for="credit_no" >&nbsp;</label>
                                                    <input type="text" id="crd04" name="crd04" class="form-control" maxlength="4" onKeyup="autotab(this, document.frmOnlineBooking.credittype)" onkeypress="return event.charCode >= 48 && event.charCode <= 57">
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                          
                                    <div class="col-sm-6"> 
                                        <div class="row">

                                            <div class="col-sm-5">
                                                <div class="form-group text-form">
                                                   <label for="credittype" >*<%=dsGeneral.item("creditcard")%></label>
                                                   <select name="credittype" id="credittype" class="form-control" >
                                                        <option value='Visa' ><%=dsGeneral.item("visa")%></option>
                                                        <option value='MasterCard' ><%=dsGeneral.item("mastercard")%></option>  
                                                        <option value='AMEX' >American Express</option>
                                                        <option value='JCB' >JCB</option>                                              
                                                   </select>
                                                </div>
                                            </div>

                                            <div class="col-sm-3">
                                                <div class="form-group text-form">
                                                    <label for="exp_mm" >*<%=dsGeneral.item("expiry_date")%></label>
                                                   <select name="expCreditMM" id="expCreditMM" class="form-control" >
                                                        <option value='--' >--</option>
                                                        <option value='01' >01</option>
                                                        <option value='02' >02</option>
                                                        <option value='03' >03</option>
                                                        <option value='04' >04</option>
                                                        <option value='05' >05</option>
                                                        <option value='06' >06</option>
                                                        <option value='07' >07</option>
                                                        <option value='08' >08</option>
                                                        <option value='09' >09</option>
                                                        <option value='10' >10</option>
                                                        <option value='11' >11</option>
                                                        <option value='12' >12</option>
                                                   </select>
                                                </div>
                                            </div>
                                                                  
                                            <div class="col-sm-1">
                                                <div class="form-group" style="text-align: center;">
                                                    <label for="exp_mm" >&nbsp;</label><p>/</p>
                                                </div>
                                            </div>
                                                                        
                                            <div class="col-sm-3">
                                                <div class="form-group text-form">
                                                    <label for="exp_yy" >&nbsp;</label>
                                                    <select name="expCreditYY" id="expCreditYY" class="form-control">
                                                        <option value='--' >----</option>
                                                        <%          
                                                        iYear=year(Now())
                                                        For i = 0 To 15
                                                            iYear=year(Now())+i
                                                            iYear2Di=Right(iYear, 2) 
                                                            Response.write "<option value='" & iYear2Di & "' >" & iYear & "</option>"
                                                        Next
                                                        %>
                                                        </select>

                                                </div>
                                            </div>

                                        </div>
                                    </div>  
                                    
                                    <div class="col-sm-12">
                                        <div class="form-group text-form">
                                            <label for="req"><%=dsGeneral.item("special_request")%></label>
                                            <textarea class="form-control" rows="3" name="remark" id="remark"></textarea>
                                            <span style="color:green"><%=dsGeneral.item("fillin_english_only")%></span>
                                        </div>
                                    </div>
                                    <div class="clearfix"></div>
                                          
                                    <div class="col-sm-12" style="text-align: right; padding-bottom: 20px;">
                                        <div class="row">
                                            <div class="col-sm-7">&nbsp;</div> 

                                            <div class="col-sm-3" style="text-align: right; padding-bottom: 20px;">
                                                <input type="button" id="btnBooking" name="btnBooking" value="Go to booking list" class="btn btn-gray" onclick="location.reload();">
                                            </div>

                                            <div class="col-sm-2" style="text-align: right; padding-bottom: 20px;">
                                                <input type="submit" id="btnBooking" name="btnBooking" value="<%=dsGeneral.item("confirm")%>" class="btn btn-blue">
                                            </div>
                                        </div>
                                    </div>
                                    <div class="clearfix"></div>   
                                    
                                </div>
                            </section>

                            <div class="clearfix"></div>
                        </div>
                        <div class="clearfix"></div>
                    <!--Guest Information-->

                </form>    
                

                <div class="clearfix"></div>



                <!--Term & Condition for Voucher Page-->
                <div class="termtxt">
                    <div style="padding:20px;font-size:14px;"> 
                        <%=dsGeneral.item("term_hmc")%>
                    </div>
                    <div class="clearfix"></div>
                </div>
                <!--Term & Condition for Voucher Page-->
            </div><!--row-->
        </div>
    </section>

    <!--#include file ="inc_file/inc_footer_main.asp"-->


    <!--send-Script parse info by xml-->
        <script>

            function autotab(current,to){
                if (current.getAttribute && 
                  current.value.length==current.getAttribute("maxlength")) {
                    to.focus() 
                    }
            }
        /*
            function autotab(original,destination){
            if (original.getAttribute&&original.value.length==original.getAttribute("maxlength"))
            destination.focus();
            }
        */

            var xmlhttp;
            function loadXMLDoc(url,cfunc)
            {
                var property = document.getElementById("property").value;
                var roomtype = document.getElementById("roomtype").value;
                var date_arrival = document.getElementById("arrival").value;
                var date_departure = document.getElementById("departure").value;

                if (window.XMLHttpRequest)
                {// code for IE7+, Firefox, Chrome, Opera, Safari
                    xmlhttp=new XMLHttpRequest();
                  }else{// code for IE6, IE5
                    xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
                }
                xmlhttp.onreadystatechange=cfunc;
                xmlhttp.open("POST",url,true);
                xmlhttp.setRequestHeader("Content-type","application/x-www-form-urlencoded");
                xmlhttp.send("property=" + property + "&roomtype=" + roomtype + "&date_arrival=" + date_arrival + "&date_departure=" + date_departure);
            }
            function loadXMLDocConfirm(url,cfunc)
            {
                var property = document.getElementById("property").value;
                var roomtype = document.getElementById("roomtype").value;
                var arrival = document.getElementById("arrival").value;
                var departure = document.getElementById("departure").value;
                var offercode = document.getElementById("offercode").value;
                var card_code = document.getElementById("card_code").value;
                var adult = document.getElementById("adult").value;
                var child = document.getElementById("child").value;
                var title = document.getElementById("title").value;
                var firstname = document.getElementById("firstname").value;
                var lastname = document.getElementById("lastname").value;
                var email = document.getElementById("email").value;
                var phone = document.getElementById("phone").value;

                var crd01 = document.getElementById("crd01").value;
                var crd02 = document.getElementById("crd02").value;
                var crd03 = document.getElementById("crd03").value;
                var crd04 = document.getElementById("crd04").value;
                var creditno = crd01+crd02+crd03+crd04;

                var expCreditMM = document.getElementById("expCreditMM").value;
                var expCreditYY = document.getElementById("expCreditYY").value;

                var credittype = document.getElementById("credittype").value;
                

                var remark = document.getElementById("remark").value;
                var prop_code = document.getElementById("prop_code").value;
                
                /*HMC more variable*/
                var membership_no = document.getElementById("membership_no").value;
                var expiry_date = document.getElementById("expiry_date").value;
                var redeem_type = document.getElementById("redeem_type").value;
                var mobile_code = document.getElementById("mobile_code").value;
                var hmc_code = document.getElementById("hmc_code").value;

                if (window.XMLHttpRequest)
                {// code for IE7+, Firefox, Chrome, Opera, Safari
                    xmlhttp=new XMLHttpRequest();
                  }else{// code for IE6, IE5
                    xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
                }
                xmlhttp.onreadystatechange=cfunc;
                xmlhttp.open("POST",url,true);
                xmlhttp.setRequestHeader("Content-type","application/x-www-form-urlencoded");

                xmlhttp.send("property=" + property + "&roomtype=" + roomtype + "&arrival=" + arrival + "&departure=" + departure + "&offercode=" + offercode+ "&card_code=" + card_code + "&adult=" + adult + "&child=" + child + "&title=" + title + "&firstname=" + firstname + "&lastname=" + lastname + "&email=" + email + "&phone=" + phone + "&creditno=" + creditno + "&expCreditMM=" + expCreditMM + "&expCreditYY=" + expCreditYY + "&remark=" + remark + "&prop_code=" + prop_code+ "&crd01=" + crd01+ "&crd02=" + crd02+ "&crd03=" + crd03+ "&crd04=" + crd04+ "&membership_no=" + membership_no+ "&expiry_date=" + expiry_date+ "&redeem_type=" + redeem_type+ "&mobile_code=" + mobile_code+ "&hmc_code=" + hmc_code);
            }

            function parseScript(strcode) {
              var scripts = new Array();         // Array which will store the script's code

              // Strip out tags
              while(strcode.indexOf("<script") > -1 || strcode.indexOf("</script") > -1) {
                var s = strcode.indexOf("<script");
                var s_e = strcode.indexOf(">", s);
                var e = strcode.indexOf("</script", s);
                var e_e = strcode.indexOf(">", e);

                // Add to scripts array
                scripts.push(strcode.substring(s_e+1, e));
                // Strip from strcode
                strcode = strcode.substring(0, s) + strcode.substring(e_e+1);
              }

              // Loop through every script collected and eval it
              for(var i=0; i<scripts.length; i++) {
                try {
                  jQuery.globalEval(scripts[i]);//eval(scripts[i]);
                }
                catch(ex) {
                  //alert(scripts[i]);// do what you want here when a script fails
                }
              }
            }

            function loadRoomType(url,cfunc)
            {
                 var property = document.getElementById("property").value;

                if (window.XMLHttpRequest)
                {// code for IE7+, Firefox, Chrome, Opera, Safari
                    xmlhttp=new XMLHttpRequest();
                  }else{// code for IE6, IE5
                    xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
                }
                xmlhttp.onreadystatechange=cfunc;
                xmlhttp.open("POST",url,true);
                xmlhttp.setRequestHeader("Content-type","application/x-www-form-urlencoded");
                xmlhttp.send("property=" + property);
            }
        </script>
    <!--send-Script parse info by xml-->

    <!--ourfunctions-Script parse info by xml-->
        <script>
                $(document).ready(function() {
                    showDeparture();	
                });

                function funChkAvi(frm)
                {
                    document.getElementById('sale-form-matrix').innerHTML ="<div class='container'><div class='col-md-1'>Loading...</div></div>";

                    loadXMLDoc("../pms-room-chkavai.asp",function()
                      {
                      if (xmlhttp.readyState==4 && xmlhttp.status==200)
                        {
                            document.getElementById("sale-form-matrix").innerHTML=xmlhttp.responseText;
                            parseScript(document.getElementById('sale-form-matrix').innerHTML);
                        }
                      });
                }

                function showRoomtype(frm)
                {
                    document.getElementById('show-roomtype').innerHTML ="<div class='container'><div class='col-md-1'>Loading...</div></div>";

                    loadRoomType("../inc_file/inc_roomtype.asp",function()
                      {
                      if (xmlhttp.readyState==4 && xmlhttp.status==200)
                        {
                            document.getElementById("show-roomtype").innerHTML=xmlhttp.responseText;
                            parseScript(document.getElementById('show-roomtype').innerHTML);
                        }
                      });
                }

                function showDeparture(frm)
                {

                    var date_arrival = document.getElementById("arrival").value;
                    var parts =date_arrival.split('/');
                    var someDate = new Date(parts[2]+"-"+parts[1]+"-"+parts[0]);
                    someDate.setDate(someDate.getDate() + 1);
                    var date_departure = someDate.toISOString().substr(0,10);
                    var partsDeparture =date_departure.split('-');

                    document.getElementById("departure").value = partsDeparture[2]+"/"+partsDeparture[1]+"/"+partsDeparture[0];
                }
        </script>
    <!--ourfunctions-Script parse info by xml-->

    <!--calendar-->
        <script src="<%=usepath%>dist/plugin/date-picker-v3/js/bootstrap-datepicker.min.js"></script>

        <script>
            function minDate()
            {
                var today = new Date();
                var dd = today.getDate()+1;
                var mm = today.getMonth()+1;//January is 0!
                var yyyy = today.getFullYear();
                if(dd < 10)
                {
                dd = '0' + dd;
                }
                if(mm < 10)
                {
                mm = '0' + mm;
                }
                var set_today = dd+"/"+mm+"/"+yyyy;
                return set_today;
            }

            function maxDate()
            {
                var today = new Date();
                var dd = today.getDate()+30;
                var mm = today.getMonth()+1;//January is 0!
                var yyyy = today.getFullYear();
                if(dd < 10)
                {
                dd = '0' + dd;
                }
                if(mm < 10)
                {
                mm = '0' + mm;
                }
                var set_today = dd+"/"+mm+"/"+yyyy;
                return set_today;
            }

            $(document).ready(function() {
                $('#arrival').datepicker({
                    startDate: minDate(),
                    //endDate: maxDate(),
                    autoSize: true,
                    orientation: "bottom left",
                    autoclose: true,
                    language: "en-GB",
                    format: "dd/mm/yyyy"
                });
                $('#expiry_date').datepicker({
                    startDate: minDate(),
                    autoSize: true,
                    orientation: "bottom left",
                    autoclose: true,
                    language: "en-GB",
                    format: "dd/mm/yyyy"
                });
            });
        </script>
    <!--calendar-->    
    
    <!--validate form-->
        <script>
            function validateForm() {
                var property = document.forms["frmOnlineBooking"]["property"].value;
                var roomtype = document.forms["frmOnlineBooking"]["roomtype"].value;
                var firstname = document.forms["frmOnlineBooking"]["firstname"].value;
                var lastname = document.forms["frmOnlineBooking"]["lastname"].value;
                var email = document.forms["frmOnlineBooking"]["email"].value;
                var crd01 = document.forms["frmOnlineBooking"]["crd01"].value;
                var crd02 = document.forms["frmOnlineBooking"]["crd02"].value;
                var crd03 = document.forms["frmOnlineBooking"]["crd03"].value;
                var crd04 = document.forms["frmOnlineBooking"]["crd04"].value;
                var expCreditMM = document.forms["frmOnlineBooking"]["expCreditMM"].value;
                var expCreditYY = document.forms["frmOnlineBooking"]["expCreditYY"].value;

                if (property == null || property == "") {
                    alert("Please select the hotels.");
                    document.forms["frmOnlineBooking"]["property"].focus();
                    return false;

                }else if(roomtype == null || roomtype == ""){
                    alert("Please select the hotel room type.");
                    document.forms["frmOnlineBooking"]["roomtype"].focus();
                    return false;
                    
                }else if(firstname == null || firstname == ""){
                    alert("First Name must be filled out.");
                    document.forms["frmOnlineBooking"]["firstname"].focus();
                    return false;
                    
                }else if(lastname == null || lastname == ""){
                    alert("Last Name must be filled out.");
                    document.forms["frmOnlineBooking"]["lastname"].focus();
                    return false;
                    
                }else if(email == null || email == ""){
                    alert("E-mail must be filled out.");
                    document.forms["frmOnlineBooking"]["email"].focus();
                    return false;

                }else if(crd01 == null || crd01 == ""){
                    alert("Credit No. must be filled out");
                    document.forms["frmOnlineBooking"]["crd01"].focus();
                    return false;
                    
                }else if(crd02 == null || crd02 == ""){
                    alert("Credit No. must be filled out");
                    document.forms["frmOnlineBooking"]["crd02"].focus();
                    return false; 
                }else if(crd03 == null || crd03 == ""){
                    alert("Credit No. must be filled out");
                    document.forms["frmOnlineBooking"]["crd03"].focus();
                    return false; 
                }else if(crd04 == null || crd04 == ""){
                    alert("Credit No. must be filled out");
                    document.forms["frmOnlineBooking"]["crd04"].focus();
                    return false; 

                }else if(expCreditMM == null || expCreditMM == "--"){
                    alert("Expiry date must be filled out");
                    document.forms["frmOnlineBooking"]["expCreditMM"].focus();
                    return false;

                }else if(expCreditYY == null || expCreditYY == "--"){
                    alert("Expiry date must be filled out");
                    document.forms["frmOnlineBooking"]["expCreditYY"].focus();
                    return false;
                
                }else{
                    return true;
                }
                   
            }
        </script>


    <!--validate form-->

</body>
</html>
<%
Else
    response.Redirect("https://www.centarahotelsresorts.com/hmc/")
End If
%>