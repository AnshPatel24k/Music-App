dim msg, sapi
msg=inputbox("Enter the text here","Quick Speech")
set sapi=createobject("sapi.spvoice")
sapi.speak msg