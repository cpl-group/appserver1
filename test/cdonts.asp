<%option explicit
dim mail
set mail = Server.CreateObject("CDONTS.newMail")

mail.Body = "Form1.Text1.Text!!!!!"
mail.To = "daniel_lasaga@genergy.com"
mail.Send
%>

sssss