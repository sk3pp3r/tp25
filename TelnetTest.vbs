Dim inpustservername,inputmailfrom,inputrcptto
inpustservername=inputBox("Enter your mailserver name or ip","MailServer") 
inputmailfrom=inputBox("Enter MAIL from","MailFrom")
inputrcptto=inputBox("Enter your mail RCPT TO","RCP TO")
Set oShell = CreateObject("WScript.Shell")
    'Start up command prompt
    oShell.run"cmd.exe"
    WScript.Sleep 500
    'Send keys to active window; change the 
    '     SERVER MAIL SMTP NAME.
    oShell.SendKeys"telnet "+ inpustservername+" 25"
    'Emulate the enter key
    oShell.SendKeys("{Enter}")
    WScript.Sleep 1000
    'write helo 
    oShell.SendKeys"helo"
    oShell.SendKeys("{Enter}")
     WScript.Sleep 500
    'write the MAIL ADDRES 
    oShell.SendKeys"mail from:"+inputmailfrom
    oShell.SendKeys("{Enter}")    
    WScript.Sleep 500
    'write write maile MAIL_ADDRESS
    oShell.SendKeys"Rcpt to:"+inputrcptto
    oShell.SendKeys("{Enter}")
     WScript.Sleep 500
     'write bye
    oShell.SendKeys"data"   
    oShell.SendKeys("{Enter}")    
     oShell.SendKeys"This is a test mail"
    oShell.SendKeys("{Enter}") 
   oShell.SendKeys"." 
 oShell.SendKeys("{Enter}")
    WScript.Sleep 500
'write to the quit
   oShell.SendKeys"quit"   
   oShell.SendKeys("{Enter}")    
     WScript.Sleep 500
     'write the password to the quit
  oShell.SendKeys"exit"   
   oShell.SendKeys("{Enter}")    
     WScript.Sleep 400
     
'Exit the program
   ' oShell.SendKeys"% "

