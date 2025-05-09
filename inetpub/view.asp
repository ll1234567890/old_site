<%@ CodePage=65001 EnableSessionState=False Language=JScript %>
<% Response.Expires=-1 %>
<%
	Response.CodePage = 65001
	Response.CharSet = "utf-8"
 
	var ymonth;
	var rs;
	var rs2;	
	var status = 0; //success
	
	try {
		rs = Server.CreateObject("ADODB.Recordset");
		rs.open("select yy.F1,yy.F2,yid.F2 as F6,yid2.F2 as F7,yy.F8,yy.F9 from y.txt yy,yid.txt yid,yid2.txt yid2 where yid.F1=yy.F6 and yid2.F1=yy.F7 and yy.F"+ Request.QueryString("act") + "=" + Request.QueryString("name") + " order by yy.F1 desc,yy.F2 desc" , "DSN=trans");
	
		var i = 1;
		while (!rs.EOF){
			var ydate = new Date(rs.Fields("F1").Value);
			ymonth = ydate.getMonth() + 1;
			Response.write ("&ydate" + i + "=")
			Response.write ( ydate.getDate() +"/"+ ymonth  +"/"+ ydate.getFullYear());
			var ytime = new Date(rs.Fields("F2").Value);
			Response.write ("&ytime" + i + "=")
			Response.write ( ytime.getHours() +":"+ ytime.getMinutes()  +":"+ ytime.getSeconds());
			Response.write ("&ydueto" + i + "=" + rs.Fields("F6") + "&yduefrom" + i + "=" + rs.Fields("F7") + "&yamount" + i + "=" + rs.Fields("F8") + "&yneed" + i + "=" + rs.Fields("F9"));

			rs.MoveNext();
			i++;
		}

		rs2 = Server.CreateObject("ADODB.Recordset");
		rs2.open("select count(F1) from y.txt where F"+ Request.QueryString("act") + "=" + Request.QueryString("name") ,"DSN=trans");
		Response.write ("&COUNT=" + rs2.Fields(0));

	}
	catch (e) {
	 	status = 101; //unhandled error;
	}



	finally {
		Response.Write("&STATUS=" + status);
		try {
			rs.close();
			rs2.close();
		} 
		catch (e) {}
		delete rs;
		rs = null;
		delete rs2;
		rs2 = null;
		
	}

%>