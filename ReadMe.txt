TITLE : SSL eMail Sender

RATE  : Advanced++
ARG   : Secure Socket Layer SMTP connection
DESC  : Sending eMail by SSL, MIME attachments and BCC comma separated
TEST  : TESTED ON BLU HOTMAIL SERVER
IDE   : Visual Basic 6 SP6
AUTHOR: GioRock
EMAIL : giorock@libero.it
SITE  : http://digilander.libero.it/giorock (Under construction - ITA)
DATE  : 24/03/2009

UPDATED 2012:
	Now program is updated to communicate about SSL3.0 Servers
	during handshak messages, Servers must be able to fallback
	at SSL2.0 protocol, otherwise it doesn't work correctly.

SUPPORTED CRYTOGRAPHY:

	128-bit RC4 with MD5 Hush
	base64
	UUEncoder

THE ONLY POSSIBLE TESTED CONFIGURATION TO WORK FASTER:

	SMTP: smtp.live.com
	PORT: 587
	USER: xxxxxx@live.it or xxxxxx@live.com or xxxxxx@hotmail.com
	PWD : ******
	SSL : TRUE
	AUTH: TRUE

   xxxxxx = your account name
   ****** = your password
   SSL    = Need a protected connection
   AUTH   = Need a Server authentication

   PS.      For other config you can change the code to adapt program

NOTE  :

Since I never found something about this VB arguments on Internet, I've
tryed to merge only 2 out of date "SSL https://... samples" discovered
on PSC in all section.
The result take me away about 3 days of hardest full immersion work to
decipher all routines and get the exact sequence of commands to send
by a SSL connection before obtain server access and open a secure 
connection to my eMails (you must issue a STARTTLS command before ...).
More things are still TODO, type of re-create the session in a new
connection, trap all errors, etc... but important is that you can send
out some messages and they run exactly where you want.
Now I can almost ensure that all procedures working out fine, Enjoy!!! 

WARNING:

-----------------------------------------------------------------------
Do not use this program to SPAM messages or other abuse sending eMails
-----------------------------------------------------------------------

SPECIFIC:

SSL is based on public key cryptography, it works in the following
manner:

C = Client
S = Server

client-hello         C -> S: challenge, cipher_specs
server-hello         S -> C: connection-id, server_certificate, 
                             cipher_specs
client-master-key    C -> S: {master_key}server_public_key
client-finish        C -> S: {connection-id}client_write_key
server-verify        S -> C: {challenge}server_write_key
server-finish        S -> C: {new_session_id}server_write_key

First the Client sends some random data known as the CHALLENGE, along
with a list of ciphers it can use, for simplicity we will only use
128-bit RC4 with MD5
The Server responds with a random data, known as the CONNECTION-ID, and
the Server's Certificate and list of cipher specs
The Client extracts the Public Key from the Server's Certificate then
uses it to Encrypt a randomly generated Master Key, this Key then sent
to the Server
The Client and Server both generate 2 keys each by hashing the Master
Key with other values, and the client sends a finish message, encrypted
with the client write key
The Server Responds by returning the CHALLENGE encrypted using the 
Client Read Key, this proves to the Clinet that the Server is who it 
says its is
The Server sends its finish message, which consists of a randomly 
generated value, this value can be used to re-create the session in a 
new connection, but that is not supported in this example

You can take a look on Wikipedia for an exaustive explanation about SSL
connections

						  GioRock 2009
