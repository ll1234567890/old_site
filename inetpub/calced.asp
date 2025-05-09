<%@ CodePage=65001 EnableSessionState=False Language=JScript %>
<% Response.Expires=-1 %>
<%
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
 

	var rs;
	var status = 0; //success
	
	try {
		rs = Server.CreateObject("ADODB.Recordset");
		rs.open("select * from ydb.txt" ,"DSN=trans");
	
		var i = 1;
		while (!rs.EOF){
			Response.write ("&a" + i + "=" + rs.Fields("F1") + "&b" + i + "=" + rs.Fields("F2") + "&c" + i + "=" + rs.Fields("F3") + "&d" + i + "=" + rs.Fields("F4") + "&e" + i + "=" + rs.Fields("F5") + "&f" + i + "=" + rs.Fields("F6"));
			rs.MoveNext();
			i++;
		}
	}
	catch (e) {
	 	status = 101; //unhandled error;
	}



	finally {
		Response.Write("&COUNT=6");
		Response.Write("&STATUS=" + status);
		try {
			rs.close();
		} 
		catch (e) {}
		delete rs;
		rs = null;
		
	}

%>