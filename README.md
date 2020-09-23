<div align="center">

## How to send email from ASP \(Part I of II\)


</div>

### Description

Want to send SMTP email from your ASP app? CDONTS is a free e-mail component that comes with IIS4, that (despite its weird name) is very easy to use and has good performance.

The setup actually takes longer than the trivial scripting. To set it up, first, make sure you have installed the SMTP service. It is installed with IIS by default, but if you've messed with your settings, you may ahve to reinstall it. Check that the SMTP service shows up in the services part of the control panel and that the file CDONTS.DLL shows up in your System32 directory.

Then using the following code. Don't forget to substitute the email address you want to send to for someaddress@someplace.com, which appears twice in the code).
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Ippolito \(vWorker\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-ippolito-vworker.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Internet/ Browsers/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-browsers-html__4-9.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-ippolito-vworker-how-to-send-email-from-asp-part-i-of-ii__4-33/archive/master.zip)





### Source Code

```
<%
Dim objMail
Set objMail = CreateObject("CDONTS.NewMail")
objMail.Send "someaddress@someplace.com","someaddress@someplace.com","Subject: You are now email enabled!","Yes, you now have the power of SMTP email from your apps, thanks to IIS and CDONTS.Newmail!"
%>
```

