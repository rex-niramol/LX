<!--#include file ="conf/inc_config.asp"-->
<!DOCTYPE html>
<!--[if IE 7 ]><html class="ie ie7" xml:lang="en-us" lang="en-us> <![endif]-->
<!--[if IE 8 ]><html class="ie ie8" xml:lang="en-us" lang="en-us"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><!--><html class="no-js" xml:lang="en-us" lang="en-us"><!--<![endif]-->
<!-- html header -->

<head>

<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Promark</title>
<meta name="description" content="">
<meta name="subject" CONTENT="">
<meta name="author" content="">
<meta name="copyright" content="">
<meta name="distribution" content="global">
<meta name="robots" content="all, index, follow">
<meta name="rating" content="general">
<link rel="author" href=""/>
<link rel="publisher" href=""/>
<link rel="shortcut icon" href="http://cdn.centarahotelsresorts.com/images/img/favicon.ico">
<!-- hreflang -->

<link rel="alternate" href="http://vip.centarahotelsresorts.com/promark/" hreflang="x-default" />
<link rel="alternate" href="http://vip.centarahotelsresorts.com/promark/" hreflang="en" />

<!-- facebook -->
<meta property="og:title" content="" />
<meta property="og:type" content="article" />
<meta property="og:url" content="vip://www.centarahotelsresorts.com/promark/" />
<meta property="og:image" content="http://www.centarahotelsresorts.com/corporateweb/dist/images/logo-centara-hotels-resorts-200x200.jpg" />
<meta property="og:description" content="" />
<!-- twitter -->
<meta name="twitter:card" content="summary" />
<meta name="twitter:site" content="#explore_chr" />
<meta name="twitter:title" content="">
<meta name="twitter:description" content=""><!--twitter description less than 200 characters-->
<meta name="twitter:creator" content="#explore_chr" />
<meta name="twitter:image" content="http://www.centarahotelsresorts.com/corporateweb/dist/images/logo-centara-hotels-resorts-280x150.jpg" /><!-- size 280x150px -->

<!-- canonical -->
<link rel="canonical" href="http://vip.centarahotelsresorts.com/promark/" />
<meta name="baidu-site-verification" content="dia3VA7sMO" />

<!-- Style -->
<link  href="assets/bootstrap/css/bootstrap.min.css" rel="stylesheet" type="text/css">
<link href="https://fonts.googleapis.com/css?family=Open+Sans:300,400,600|Roboto:400,500,700" rel="stylesheet">
<link href="http://fontawesome.io/assets/font-awesome/css/font-awesome.css" rel="stylesheet">
<link href="dist/plugin/masterslider/css/masterslider.css" rel="stylesheet" type="text/css">
<link href="dist/plugin/masterslider/css/style.css" rel="stylesheet" type="text/css">
<link href="dist/plugin/date-picker-v3/css/bootstrap-datepicker.min.css" rel="stylesheet" type="text/css">
<link href="dist/style/desktop/main-css.css" rel="stylesheet" type="text/css">
<link href="dist/css/promark.css" rel="stylesheet" type="text/css">
<!--HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
<!--[if lt IE 9]>
<script  src="assets/bootstrap/js/html5/html5shiv.min.js"></script>
<script src="assets/bootstrap/js/html5/respond.min.js"></script>
<![endif]-->
<script src="assets/jquery/jquery-1.11.1.min.js"></script>
<script src="assets/bootstrap/js/bootstrap.min.js"></script>
<script src="dist/plugin/masterslider/js/masterslider.min.js"></script>
<script src="dist/plugin/masterslider/js/jquery.easing.min.js"></script>

<!-- Google Tag Manager by Centara Hotels & Resorts -->
<noscript><iframe src="//www.googletagmanager.com/ns.html?id=GTM-9N5X"
height="0" width="0" style="position:absolute; top:0; display:none;visibility:hidden;"></iframe></noscript>
<script>(function(w,d,s,l,i){w[l]=w[l]||[];w[l].push({'gtm.start':
new Date().getTime(),event:'gtm.js'});var f=d.getElementsByTagName(s)[0],
j=d.createElement(s),dl=l!='dataLayer'?'&l='+l:'';j.async=true;j.src=
'//www.googletagmanager.com/gtm.js?id='+i+dl;f.parentNode.insertBefore(j,f);
})(window,document,'script','dataLayer','GTM-9N5X');</script>
<!-- End Google Tag Manager by Centara Hotels & Resorts -->

</head>

<body>



<style>
label{font-size:14px; color:#111; font-family: 'Montserrat', sans-serif; text-transform:uppercase;}
.form-group{margin-bottom:20px;}
.wrapper {
    padding: 0px 0px;
}
</style>
<section class="wrapper white">
	<div class="container">
    	<div class="row">

<div class="container">
	<div class="col-md-offset-3 col-md-6" style="background:#fff; margin-top:30px; margin-bottom:30px;">
    	<div class="row">
    	<div style="padding-top:15px; padding-bottom:15px;">
            <div class="col-sm-12 text-center">
            	<img src="<%=usepath%>dist/c1c/c1c-logo.png">
                <hr>
            </div>
            <div class="clearfix"></div>
            
            <div class="col-sm-12 text-center">
            <h4><%=dsC1CForm.item("we_promise_desc")%></h4>
			<% If SESSION("access_key")="" Then
				Response.write "<h4><font color='red'>" & SESSION("message") & "</font></h4>"
			   End If
			%>
            </div>
            <div class="clearfix"></div>
            
            <div class="col-md-offset-2 col-md-8" style="margin-top:30px;">
            
                <form method="post" action="http://www.centarahotelsresorts.com/WEBSERVICE/API/CHR_API_AUTHEN_POST.asp">
                  <div class="form-group">
                    <label for="exampleInputEmail1"><%=dsC1CForm.item("email")%>*</label>
                    <input type="email" id="member_username" name="member_username" class="form-control input-sm">
                  </div>
                  <div class="form-group">
                    <label for="exampleInputPassword1"><%=dsC1CForm.item("password")%>*</label>
                    <input type="password" id="member_userpsw" name="member_userpsw" class="form-control input-sm">
                  </div>
				<input type="hidden" id="api_username" name="api_username"  value="centara">
				<input type="hidden" id="member_result" name="member_result" value="1">
				<input type="hidden" name="page" id="page" value="promark/login_c1c_check.asp">
                  <div class="checkbox">
                    <label>
                      <input type="checkbox" id="member_keeploged" name="member_keeploged"> <%=dsC1CForm.item("remember_me")%>
                    </label>
                  </div>
                  <button type="submit" class="btn btn-red"><%=dsC1CForm.item("sign_in")%></button>
                  <a href="http://centarathe1card.com/forget/" class="btn btn-link" target="_blank"><small><%=dsC1CForm.item("lost_your_password")%></small></a>
                </form>            
            
            </div>
            <div class="clearfix"></div>
            
            
            
        </div>
        
        <div style="background:#f1f1f2; padding-top:15px; padding-bottom:15px;">
        <div class="col-sm-12 text-center"><h5><%=dsC1CForm.item("not_a_member")%></h5><a href="http://www.centarahotelsresorts.com/sign-up/" class=" btn btn-default btn-sm"><%=dsC1CForm.item("sign_up")%></a> </div>
        <div class="clearfix"></div>
        </div>
        

            
            </div>
            </div>
            </div>
			<div class="clearfix"></div>
		<br>
        </div>
    </div>
</section>


</body>
</html>
