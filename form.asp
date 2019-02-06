
<!--#include file ="conf/inc_config.asp"-->

<head>
<!--#include file ="../Phase 1/inc_file/inc_metatag.asp"-->
<!--#include file ="../Phase 1/inc_file/inc_style.asp"-->
<!--#include file ="../Phase 1/inc_file/inc_script.asp"-->
<style>
.btn-black {
    color: #fff;
    background-color: #000000;
    border-color: #000000;}
.btn-black:hover,.btn-red:focus,.btn-red:active,.btn-red.active,.open .dropdown-toggle.btn-red{
	color:#000000;background:none;border-color:#000000;}
}
</style>
</head>

<body style="background-color:#ffffff;" >

<!--#include file ="conf/set_condition_language.asp"-->
<!--#include file ="inc_file/inc_dropdownlang.asp"-->

<%'----- Customer Info -----%>
<header class="header-container" style="z-index:9; position:relative;" >
<div class="container">

	<div class="row">
    <div class="col-sm-12" >
		<div class="col-sm-4">
    	<a href="http://www.centarahotelsresorts.com/" title="Centara Hotels & Resorts"><img src="http://www.centarahotelsresorts.com/promark/dist/images/chr-red.png" class="img-responsive"></a>
        <div class="clearfix"></div>
        </div>
        <div class="col-lg-4 col-md-4 col-sm-6 col-xs-12 text-left pull-right" style="padding-top:20px; padding-bottom:20px; color:#b5121b;">
			<div class="contact-right">

<strong><%=first_name%> </strong>| <a href="http://www.centarahotelsresorts.com/promark/logout.asp"><%=dsC1CForm.item("log_out")%></a><br/>
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


<%'----- For Mobile -----%>
<div class="header-sm"> 

<div class="col-xs-6 col-sm-12 text-center">
<a href="http://www.centarahotelsresorts.com/">
<img src="http://www.centarahotelsresorts.com/promark/dist/images/chr-red.png" class="img-responsive" style="margin:0 auto;">
</a>
</div>

<div class="clearfix"></div>

<div class="col-xs-12 text-center" style="padding-bottom:20px;">

<strong><%=first_name%> </strong>| <a href="http://www.centarahotelsresorts.com/th/promark/logout.asp">LOG OUT</a><br/>
<strong>ID : <%=guest_id%><br/>
<span style="color:#857650;"><%=card_type%></span></strong><br/>
<strong><span style="color:#000000;"><%=prop_name%></span></strong>

</div>

<div class="clearfix"></div>

</div>


<%'----- Banner -----%>
<section id="banner" style="max-height:440px; overflow:hidden;">
		<img src="http://www.centarahotelsresorts.com/promark/dist/images/banner-gold.jpg" alt="Gold Card"/ class="img-responsive" style="margin:0 auto;">
<div class="clearfix"></div>
</section>

<%'----- Card Type Line -----%>
<section class="headline" >
 		<img src="<%=card_title%>" alt="Gold Card"/ class="img-responsive" style="margin:0 auto;">
<div class="clearfix"></div>
</section>



<%'-- Section Show Voucher --%>
<section class="bodytxt">
<div class="container">
<div class="row">
 	
    <h3 style="color:#1c1b19;"><%=dsGeneral.item("use_your_voucher")%></h3> 
    
<div class="col-sm-12 contentbox">



<div class="clearfix"></div>

<%'----- Term & Condition for Voucher Page ----- %>
<div class="termtxt">
<div class="" style="padding:20px;"> 

<%=dsGeneral.item("term_pm")%>	

<div class="clearfix"></div>
</div>
</div>
<div class="clearfix"></div>

</div>

<div class="clearfix"></div>
</div>
</div>
</section>

<!--#include file ="inc_file/inc_footer_main.asp"-->
<!--#include file ="inc_file/inc_remarketing.asp"-->
</body>
</html>
