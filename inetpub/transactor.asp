<%
function WriteToFile(FileName, Contents, Append)
on error resume next

if Append = true then
   iMode = 8
else 
   iMode = 2
end if
set oFs = server.createobject("Scripting.FileSystemObject")
set oTextFile = oFs.OpenTextFile(FileName, iMode, True)
oTextFile.Write Contents
oTextFile.Close
set oTextFile = nothing
set oFS = nothing

end function

%>

<HTML>
<BODY>

<%

'WriteToFile "C:\Inetpub\wwwroot\transactor\Test2.txt", chr(13) & chr(10) & "I am fine", false
'dim todaysDate
'todaysDate=now()


'response.write(Replace(time(), ":", "-") & chr(13) & chr(10) & chr(13)) 





WriteToFile "C:\Inetpub\wwwroot\transactor\y" & Replace(Request.Servervariables("REMOTE_ADDR"),".","_") & "_" & Replace(time(), ":", "-") & ".txt", Request.QueryString("ydata"), false


%>
&STATUS=0&
</BODY>
</HTML>