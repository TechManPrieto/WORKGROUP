message=InputBox("What do you want me to say?","Windows Text To Speech") 
Set sapi=CreateObject("sapi.spvoice") 
sapi.Speak message