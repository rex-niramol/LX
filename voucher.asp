<%
'----- Check Login -----
If Session("promark_login") = True Then

'----- Get Guest ID -----
guest_id =  Session("guest_number")

%>

<!--#include file ="conf/inc_config.asp"-->

<%

setdropdownlang = True
arrLangSet = array("th","en")

redeem = dsGeneral.item("redeem")
term = dsSpecial.item("terms_and_conditions")
c1c_login = dsGeneral.item("c1c_login")


'response.write check_guest

'----- Query Customer Info -----

strSQL3 = ""
strSQL3 = strSQL3 & "Select top 1 pm_tb_r_sales.first_name,pm_tb_r_sales.card_type,properties.prop_name,properties.prop_name_l_TH,properties.bookingcode,pm_tb_m_properties.prop_booking_id,pm_tb_m_properties.prop_code FROM( "
strSQL3 = strSQL3 & "pm_tb_r_sales LEFT JOIN pm_tb_m_properties ON pm_tb_r_sales.prop_code = pm_tb_m_properties.prop_code "
strSQL3 = strSQL3 & "LEFT JOIN properties ON properties.prop_code = pm_tb_r_sales.prop_code "
strSQL3 = strSQL3 & ")"
strSQL3 = strSQL3 & "Where pm_tb_r_sales.guest_id = '"& guest_id &"' "
Set objRec3 = Server.CreateObject("ADODB.Recordset")
objRec3.Open strSQL3, Conn, 1,3


If Not objRec3.EOF then
		
	first_name = objRec3("first_name")
	card_type = objRec3("card_type")
	prop_code = objRec3("prop_code")
	prop_name = objRec3("prop_name")
	bookingcode = objRec3("bookingcode")
	prop_booking_id = objRec3("prop_booking_id")
	
	If session("admin_lang_code") = "TH" Then
	prop_name = objRec3("prop_name_l_TH")
	pm_path = webpath&"th/ultimate-dining/"
	lang_code = "th_TH"
	Else
	prop_name = objRec3("prop_name")
	pm_path = webpath&"ultimate-dining/"
	lang_code = "en_EN"
	End If
	
	'----- Check Card Type -----
	If card_type = "gold" then
		card_type = dsGeneral.item("gold_card")
		card_title = "../ultimate-dining/dist/images/gold-card.png"

	Else 
		card_type = dsGeneral.item("platinum_card")
		card_title = "../ultimate-dining/dist/images/platinum-card.jpg"
	End If
	
	'----- Check Session Member ID -----
	If Session("member_id") <> "" Then

	redeemlink = pm_path&"online-booking.asp"
	
	Else
	
	redeemlink = pm_path&"online-booking.asp"
	
	End If
	
%>

<head>
<!--#include file ="inc_file/inc_metatag.asp"-->
<!--#include file ="inc_file/inc_style.asp"-->
<!--#include file ="inc_file/inc_script.asp"-->
<style>
.btn-black {
    color: #fff;
    background-color: #000000;
    border-color: #000000;}
.btn-black:hover,.btn-red:focus,.btn-red:active,.btn-red.active,.open .dropdown-toggle.btn-red{
	color:#000000;background:none;border-color:#000000;}
.term:hover {
    text-decoration: underline;
}
</style>
</head>

<body style="background-color:#ffffff;" >

<!--#include file ="conf/set_condition_language.asp"-->
<!--#include file ="inc_file/inc_dropdownlang.asp"-->


<%'------ check 5 digits right' ---------
check_guest = right(guest_id,5)
'check_guest = CInt(check_guest)
'response.write check_guest
%>

<%'----- Customer Info -----%>
<header class="header-container" style="z-index:9; position:relative;" >
<div class="container">

	<div class="row">
    <div class="col-sm-12" >
		<div class="col-sm-4">
    	<a href="<%=chr_link%>" title="Centara Hotels & Resorts"><img src="../ultimate-dining/dist/images/chr-red.png" class="img-responsive"></a>
        <div class="clearfix"></div>
        </div>
        <div class="col-lg-4 col-md-4 col-sm-6 col-xs-12 text-left pull-right" style="padding-top:20px; padding-bottom:20px; color:#b5121b;">
			<div class="contact-right">

<strong><%=first_name%> </strong>| <a href="<%=usepath%>logout.asp"><%=dsC1CForm.item("log_out")%></a><br/>
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
<a href="<%=chr_link%>">
<img src="../ultimate-dining/dist/images/chr-red.png" class="img-responsive" style="margin:0 auto;">
</a>
</div>

<div class="clearfix"></div>

<div class="col-xs-12 text-center" style="padding-bottom:20px;">

<strong><%=first_name%> </strong>| <a href="<%=usepath%>logout.asp">LOG OUT</a><br/>
<strong>ID : <%=guest_id%><br/>
<span style="color:#857650;"><%=card_type%></span></strong><br/>
<strong><span style="color:#000000;"><%=prop_name%></span></strong>

</div>

<div class="clearfix"></div>

</div>


<%'----- Banner -----%>
<section id="banner" style="max-height:440px; overflow:hidden;">
		<img src="../ultimate-dining/dist/images/banner-gold.jpg" alt="Gold Card"/ class="img-responsive" style="margin:0 auto;">
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

<%

'----- Query Check Voucher with this Guest ID -----

strSQL2 = "SELECT Distinct pm_tb_a_offers.rn_no,pm_tb_r_sales.card_type,pm_tb_m_properties.prop_name,pm_tb_r_sales.prop_code,pm_tb_a_offers.offer_code,pm_tb_a_offers.offer_name,pm_tb_a_offers.offer_ads,pm_tb_a_offers.offer_photo_thumb,pm_tb_a_offers.offer_photo_full from pm_tb_r_sales LEFT JOIN pm_tb_a_offers ON pm_tb_r_sales.prop_code = pm_tb_a_offers.prop_code LEFT JOIN pm_tb_m_properties ON pm_tb_r_sales.prop_code = pm_tb_m_properties.prop_code where pm_tb_r_sales.guest_id = '"& guest_id &"' and pm_tb_r_sales.prop_code = '"& prop_code &"' and pm_tb_a_offers.active_sts ='1' and pm_tb_a_offers.delete_sts ='N' order by pm_tb_a_offers.rn_no"

Set objRec2 = Server.CreateObject("ADODB.Recordset")
objRec2.Open strSQL2, Conn, 1,3

i=1
While Not objRec2.EOF

prop_code = objRec2("prop_code")
offer_code = objRec2("offer_code")
offer_thumb = objRec2("offer_photo_thumb")
offer_full = objRec2("offer_photo_full")
vcid = objRec2("rn_no")



' ----- Query Offer No. by Offer Code -----

strSQL = ""
strSQL = strSQL & " SELECT offer_no from pm_tb_a_offers where offer_code = '"& offer_code &"' and prop_code = '"& prop_code &"' and active_sts = '1' and delete_sts = 'N'  "
Set objRec = Server.CreateObject("ADODB.Recordset")
objRec.Open strSQL, Conn, 1,3

If Not objRec.EOF then

offer_no = objRec("offer_no") 

End If
objRec.Close()
Set objRec = Nothing


'----- Query count row -----
varShowMsg=""
strSQL = ""
strSQL = strSQL & "  Select count(guest_id) as voucher_used from pm_tb_r_booking where guest_id = '"& guest_id &"' and offer_code = '"& offer_code &"'and parent_id = '"& prop_code &"' and (sale_sts = 'C' or sale_sts = 'P') and active_sts= '1' and delete_sts = 'N' "
Set objRec = Server.CreateObject("ADODB.Recordset")
objRec.Open strSQL, Conn, 1,3

If Not objRec.EOF then

voucher_used = objRec("voucher_used") 

End If
objRec.Close()
Set objRec = Nothing

If (prop_code = "cmbr") and (check_guest > 20500) and offer_code = "cer04" then
	
		offer_thumb = "../ultimate-dining/dist/images/voucher/cmbr/buffet-new.jpg"
		offer_full = "../ultimate-dining/dist/images/voucher/cmbr/buffet-term.jpg"
			
Elseif (prop_code = "cmbr") and (check_guest <= 20500) and offer_code = "cer04" then
		offer_thumb = "../ultimate-dining/dist/images/voucher/cmbr/compli-chr.jpg"
		offer_full = "../ultimate-dining/dist/images/voucher/cmbr/compli-chr-term.jpg"

Else
offer_thumb = usepath&"dist/images/voucher/"&offer_thumb
offer_full =  usepath&"dist/images/voucher/"&offer_full
'response.write offer_thumb'

End If

check_voucher = Cint(offer_no) - Cint(voucher_used)

If Cint(check_voucher) > 0 Then 

'----- Display Voucher -----
        chkMsg="Y"
		response.write"<div class='col-sm-6 col-sm-offset-3' style='margin-bottom:20px; display:visible;'>"
 		response.write"<div class='col-sm-12 boxbordertwo' style=' background:url("& offer_thumb &") no-repeat left top; background-size: 100% 100%; border-radius:0px;'>"
		response.write"<div style='position:absolute;bottom:7px;right:30px;'>"
		response.write"<a href='"& redeemlink &"?vcid="& vcid &"' class='btn btn-sm btn-red'><strong>"&redeem&"</strong></a>"
		response.write"</div>"
		response.write"<div class='clearfix'></div>"
 		response.write"</div>"
 		response.write"<div class='col-sm-12' style='text-align:left; padding-left:0px; padding-top:10px;'>"
		response.write"<div style=' '>"
		response.write"<button type='button' class='term' style='border:none; background:none; color:#102254; font-size:16px;' id='showmodal"&i&"'><strong>"&term&"</strong></button>"
		response.write"</div>"
		response.write"<div class='clearfix'></div>"
		response.write"</div>"
		response.write"<div class='clearfix'></div>"
		response.write"</div>"
		
'----- Show modal Team & Condition -----

		response.write"<div id='myModal"&i&"' class='modal fade'>"
		response.write"<div class='modal-dialog'>"
		response.write"<div id='modalContainer' class='modal-dialog modal-lg'>"
		response.write"<div class='modal-content'>"
		response.write"<img src='"& offer_full &"' class='img-responsive' style='width:100%;' />"
		response.write"</div>"
		response.write"</div>"
		response.write"</div>"
		response.write"</div>"
		
		response.write"<script>"
		response.write"jQuery(document).ready(function(){"
		response.write"jQuery('#showmodal"&i&"').click(function(){"
		response.write"jQuery('#myModal"&i&"').modal('show');});});"
		response.write"</script>"

Else 
		varShowMsg="<h3 style='color:#1c1b19;'>No coupons available.</h3>"


End If

objRec2.MoveNext
i = i+1
Wend


objRec2.Close()
Set objRec2 = Nothing

If chkMsg="Y" Then 
     varShowMsg=""
End If
Response.write varShowMsg

%>	

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
<%
'----- For Customer info -----
End If
objRec3.Close()
Set objRec3 = Nothing

'----- Check user not login -----
Else
url2="index/"
response.Redirect(url2)
End If
%>