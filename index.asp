<!--#include file ="conf/inc_config.asp"-->
<!DOCTYPE html>

<head>
<!--#include file ="inc_file/inc_metatag.asp"-->
<!--#include file ="inc_file/inc_style.asp"-->
<!--#include file ="inc_file/inc_script.asp"-->
</head>

<body style="z-index:999; background:#000000 url(dist/images/centara-bg.png) no-repeat left top ; display:block;" >
<!--#include file ="inc_file/inc_header.asp"-->

<section style="z-index:9; width:100%; overflow:hidden; padding-top:60px; text-align:center; color:#fff;">
	<div align="center"><img src="dist/images/voucher/centara-privilege-club-white.png" alt="Card" style="height:250px"></div>
</section>

<section style="padding-top:40px;padding-bottom:20px; text-align:center;">
	<div class="container">
 		<h3 style="color:#fff;">HMC Complimentary Room Online Reservation</h3> 
 		<form method="post" class="form-horizontal" onsubmit="submit_form.asp" >

 		<div class="col-sm-12 form-group " style="text-align:center; padding-top:20px;">
  			<div class="col-sm-4" >
    		</div>
    		<div class="col-sm-4">
   	 			<input type="text" name="username" id="username" value="" class="form-control input-sm" placeholder="Username">
			 </div>
    		<div class="col-sm-4">
    		</div>
    		<div class="clearfix"></div>
		</div>
		<div class="clearfix"></div>
 		
        <div class="col-sm-12 form-group ">
  
    		<div class="col-sm-4">
    		</div>
    		<div class="col-sm-4">

            <label style="color:#fff">Password:</label>
            <input type="password" name="password" id="password" value="" class="form-control input-sm">

    
    		</div>
    		<div class="col-sm-4">
    		</div>
    		<div class="clearfix"></div>
		</div>

		<div class="clearfix"></div>

		<div class="col-sm-12 form-group">

  			<div class="col-sm-4">
    		</div>
    		<div class="col-sm-4" style="text-align:right; padding-right:30px;">
    			<div class="row">

    				<!--<button type="submit" class="btn btn-sm btn-red">LOGIN</button>
                    <a href="javascript:void(0);" onclick="form_submit();" class="btn btn-sm btn-red">Login</a>-->
                    <input type="button" value="Login" onclick="form_submit();" class="btn btn-sm btn-red">
    
    				<div class="clearfix"></div>
    			</div>
    			<div class="clearfix"></div>
    		</div>
    		
            <div class="col-sm-4">
    		</div>
   			<div class="clearfix"></div>  
		</div>

		<%
		
		'loginstatus = Session("errorstatus") 
		'response.write loginstatus
		
		
		'If loginstatus = "Error"  Then 
		'response.write "Please Login"
		'Else
		'response.write ""
		''Session("errorstatus") = 0
		'End if 

		%>
        
        <div style="width:300px;margin:auto;">
            <div id="error_message" class="topic_red_bar" style="width:300px;margin:auto;display:none;color:red;">
                Invalid your authen information, Please try again.
            </div>					
            <div style="height:3px;"></div>
        </div>
     
	</form> 
        
        
    <form method='post' name = 'online-booking' action='online-booking/'>
        <input type='submit' style='display: none;'>
    </form>
    <div class="clearfix"></div>

</div>
</section>


<script>
function clear_alert()
{
	$("#error_message").hide("slow");
	
}

function form_submit()
{
	clear_alert();
	var flag = true;
	var username = $("#username").val();
	var password  = $("#password").val();

	if(username == "")
	{
		$("#error_message").show("slow");
		flag = false;
	}


	if(flag)
	{
		$.post("submit_form.asp",
			{"username": username,"password": password},
			function(data){
				if(data == "HMC")
				{
					//alert(data);
					document.forms['online-booking'].submit();	
				}else{
					$("#error_message").show("slow");
				}
				
		});
	}
}
</script>

</body>
</html>
