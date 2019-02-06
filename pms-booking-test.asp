<%

    Response.Write "<br>START<hr>"
    Set oXmlHTTP = CreateObject("Microsoft.XMLHTTP")
    oXmlHTTP.Open "POST", "http://www.chrnetwork.com/promark/ProMark.asmx", False 

    oXmlHTTP.setRequestHeader "Content-Type", "text/xml; charset=utf-8" 
    oXmlHTTP.setRequestHeader "SOAPAction", "http://tempuri.org/saveBooking"

    SOAPRequest ="<?xml version=""1.0"" encoding=""utf-8""?>" & _
      "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
      "<soap:Body>" & _
        "<saveBooking xmlns=""http://tempuri.org/"">" &_
          "<property>ckbr</property>" &_
          "<title>MR.</title>" &_
          "<first_name>SOMCHART</first_name>" &_
          "<last_name>MITTAREE</last_name>" &_
          "<email>somchart.mittaree@gmail.com</email>" &_
          "<phone>0865366734</phone>" &_
          "<arrive>12/01/2018</arrive>" &_
          "<depart>13/01/2018</depart>" &_
          "<room_night>1</room_night>" &_
          "<room_type>DXGDK</room_type>" &_
          "<adult>2</adult>" &_
          "<child>0</child>" &_
          "<credit_no>20254165</credit_no>" &_
          "<price>2000</price>" &_
          "<business_code>cer004</business_code>" &_
          "<guest_id>5407163422</guest_id>" &_
          "<pms>COMANCHE</pms>" &_
          "<card_exp>xx/21</card_exp>" &_
          "<card_type>GLDCOM</card_type>" &_
          "<market>LOYAL</market>" &_
          "<remark>Ref.34694582 (PROMARK)</remark>" &_
          "<card_owner>CGLB</card_owner>" &_
        "</saveBooking>" &_
      "</soap:Body>" &_
      "</soap:Envelope>"

    oXmlHTTP.send SOAPRequest    
    Response.Write oXmlHTTP.responseText
    Response.Write "<br><hr>END"
%>
