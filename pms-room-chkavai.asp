<%
    prop_code = UCase(request.Form("property"))
    roomtype = UCase(request.Form("roomtype"))
    date_arrival = request.Form("date_arrival")
    date_departure = request.Form("date_departure")


    '******CommandChe Web Service '******Allotment******'
        Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
        oXmlHTTP.Open "POST", "http://www.chrnetwork.com/promark/ProMark.asmx", False 

        oXmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8" 
        oXmlHTTP.setRequestHeader "SOAPAction", "http://tempuri.org/checkAvai"

        SOAPRequest ="<?xml version=""1.0"" encoding=""utf-8""?>" & _
          "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<checkAvai xmlns=""http://tempuri.org/"">" &_
              "<property>"&prop_code&"</property>" &_
              "<arrive>"&date_arrival&"</arrive>" &_
              "<depart>"&date_departure&"</depart>" &_
              "<room_type>"&roomtype&"</room_type>" &_
            "</checkAvai>" &_
          "</soap:Body>" &_
          "</soap:Envelope>"

        oXmlHTTP.send SOAPRequest  
    '******CommandChe Web Service '******Allotment******'

    '******Send Result******'
        varResponse=oXmlHTTP.responseText
        'Response.Write=varResponse
        varSuccess=InStr(LCase(varResponse),"success")
        If varSuccess>0 Then 'Success   

            Response.write "<div  style='text-align: center; padding-bottom: 10px; border: #deecc8 solid 1px; background: #deecc8; margin: 0px 0px 20px 0px; padding-top: 10px;'>"
            Response.Write "Room Available, please key-in all information and confirm."
            Response.write "</div>"

        Else 'Fail

            Response.write "<div  style='text-align: center; padding-bottom: 10px; border: #970b0b solid 1px; background: #970b0b; margin: 0px 0px 20px 0px; padding-top: 10px;color:#ffffff;'>"
            Response.Write "No room available for this complimentary.<br>"
            Response.Write "Please do select another check-in date."
            Response.write "</div>"

        End If   
    '******Send Result******'

%>
