
<%
	Dim MyBody
    Dim MyCDONTSMail
    Set MyCDONTSMail = CreateObject("CDONTS.NewMail")
    MyCDONTSMail.From= "gulatisteel@airtelmail.in"
	MyCDONTSMail.To= "gulatisteel@airtelmail.in, gulati_nitin@hotmail.com"
	MyCDONTSMail.Bcc= "supportabacus@gmail.com"
	MyCDONTSMail.Subject="Feedback/Enquiry Information"
    MyBody = request.form("comments") & vbCrLf
				
			MyBody = MyBody &" This information is receieved from the site - Enquiry Information-: " & vbCrLf
			MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf
			MyBody = MyBody &" Name : " & request.form("sname") & vbCrLf
			MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf
			MyBody = MyBody &" Organisation : " & request.form("organisation") & vbCrLf
			MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf
			MyBody = MyBody &" E-mail : " & request.form("email") & vbCrLf  
			MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf
        	MyBody = MyBody &" Phone No. : " & request.form("phoneno") & vbCrLf  
			MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf 	
			MyBody = MyBody &" Fax No : " & request.form("Fax") & vbCrLf
        	MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf
			 	
			MyBody = MyBody &" Address : " & request.form("Address") & vbCrLf
			MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf			
			MyBody = MyBody &" City : " & request.form("city") & vbCrLf
			MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf
			MyBody = MyBody &" Pincode : " & request.form("pincode") & vbCrLf
			MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf
			MyBody = MyBody &" Country : " & request.form("country") & vbCrLf
			MyBody = MyBody &" ---------------------------------------------------------------------- " & vbCrLf 
			MyBody = MyBody &"  " & vbCrLf 
			MyBody = MyBody &" -------------------------Description/Details---------------------------- " & vbCrLf
			MyBody = MyBody &" Details of requirements : " & request.form("Description") & vbCrLf
			
    MyCDONTSMail.Body= MyBody
    MyCDONTSMail.Send
    set MyCDONTSMail=nothing
  response.redirect("thanks.html")
	
%>