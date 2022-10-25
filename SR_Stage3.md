## Stage 3 - ASP (Classic to keep things simple)
Assuming you have IIS installed on Windows and have enabled it to process ASP pages you can add the following code to the file named Batlog.asp<br>
This is required for me because I couldn't install the ODBC Driver on my work PC.

```
<% 

Dim power
power=Request.QueryString("power")
response.write(power)

    Dim con
	Set con = Server.CreateObject("ADODB.Connection")
    With con
    '   .CursorLocation = adUseClient
    '    .Mode = adModeRead
        .ConnectionString = "Driver={ODBC Driver 18 for SQL Server};Server=tcp:[your azure SQL Server].database.windows.net,1433;Database=[Database Name];Uid=[your username];Pwd=[your password];Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
        .Open
    End With
	
	con.Execute "UPDATE WORK_BATTERY SET [Percent]=" & power
	con.close
	set con = nothing
	
 %>


```
