POP3/SMTP local bridge (P/SLB)
Developed by Tomas Banovec in June 2002
e-mail: banovec@yahoo.com
        banovec@bdr.sk

==============================================================
Description:
- P/SLB is application which allows you on computers in your local
 network recieving and sending e-mails by POP3 and SMTP protocol

Features:
- this application can send mails via servers which require 
  authentification
- it is not necessary to write password into this application.
  It is required write only user name. It brings highier
  security level. Connection type (if required authentification),
  password is auromatically detected from your e-mail client.

HOW TO SET UP E-MAIL CLIENT:
- for examle outlook express: You will set your account like when 
  you are connecting directly. Only one change is address of 
  pop3/smtp server. There will not be valid smtp/pop3 server 
  address, but local IP address or name of computer, which have
  direct connection to internet. Valid IP adddress is set in 
  accounts.txt

USAGE:
- Windows 95 or better

Configuration:
- all configuration is in accounts.txt which contains this 
  informations:
username	POP3 server	SMTP server	e-mail address

All informations MUST BE splitted by TAB key. For usage see 
example of accounts.txt - one line MUST be providing only 
informations to one account. Also user name, pop3 server address,
smtp server address and e-mail address don't have to contain
spaces.


Limitations:
- only one account with same user name can be viewed 
(for example only banovec@bdr.sk can be viewed 
not banovec@yahoo.com - because of same username while logging on)

Credits:
- I'd like to thanx to www.vbip.com for their base64_encode 
  function, to Mr. Kevin Robert Keegan for his Perl SMTP AUTH
  example where I have find out how authentification works 
  (RFC's are really stupid...)and also myself to find a time 
  to code this ;)). 

