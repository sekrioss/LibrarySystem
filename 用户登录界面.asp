<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/yyy.asp" -->
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("UsersID") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("UsersID")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_yyy_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.Users WHERE UsersID = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 200, 1, 5, Recordset1__MMColParam) ' adVarChar

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
' *** Validate request to log in to this site.
MM_LoginAction = Request.ServerVariables("URL")
If Request.QueryString <> "" Then MM_LoginAction = MM_LoginAction + "?" + Server.HTMLEncode(Request.QueryString)
MM_valUsername = CStr(Request.Form("textfield"))
If MM_valUsername <> "" Then
  Dim MM_fldUserAuthorization
  Dim MM_redirectLoginSuccess
  Dim MM_redirectLoginFailed
  Dim MM_loginSQL
  Dim MM_rsUser
  Dim MM_rsUser_cmd
  
  MM_fldUserAuthorization = ""
  MM_redirectLoginSuccess = "用户主页面.asp"
  MM_redirectLoginFailed = "用户登录界面.asp"

  MM_loginSQL = "SELECT UsersID, UsersMIMA"
  If MM_fldUserAuthorization <> "" Then MM_loginSQL = MM_loginSQL & "," & MM_fldUserAuthorization
  MM_loginSQL = MM_loginSQL & " FROM dbo.Users WHERE UsersID = ? AND UsersMIMA = ?"
  Set MM_rsUser_cmd = Server.CreateObject ("ADODB.Command")
  MM_rsUser_cmd.ActiveConnection = MM_yyy_STRING
  MM_rsUser_cmd.CommandText = MM_loginSQL
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param1", 200, 1, 5, MM_valUsername) ' adVarChar
  MM_rsUser_cmd.Parameters.Append MM_rsUser_cmd.CreateParameter("param2", 200, 1, 6, Request.Form("textfield2")) ' adVarChar
  MM_rsUser_cmd.Prepared = true
  Set MM_rsUser = MM_rsUser_cmd.Execute

  If Not MM_rsUser.EOF Or Not MM_rsUser.BOF Then 
    ' username and password match - this is a valid user
    Session("MM_Username") = MM_valUsername
    If (MM_fldUserAuthorization <> "") Then
      Session("MM_UserAuthorization") = CStr(MM_rsUser.Fields.Item(MM_fldUserAuthorization).Value)
    Else
      Session("MM_UserAuthorization") = ""
    End If
    if CStr(Request.QueryString("accessdenied")) <> "" And false Then
      MM_redirectLoginSuccess = Request.QueryString("accessdenied")
    End If
    MM_rsUser.Close
    Response.Redirect(MM_redirectLoginSuccess)
  End If
  MM_rsUser.Close
  Response.Redirect(MM_redirectLoginFailed)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
body,td,th {
	font-family: "宋体";
}
</style>
</head>

<body>
<form id="form3" name="form3" method="POST" action="<%=MM_LoginAction%>">
  <img src="https://logosc.cn/uploads/output/2020/06/14/cbe769a523095d18c6c11252415cc065.jpg?t=1592122652" alt="" width="369" height="156"/>
  <table width="1000" border="1" align="center">
    <tr>
      <th height="75" align="center" bgcolor="#6699FF" scope="col">欢迎您登录用户界面</th>
    </tr>
  </table>
  <p>&nbsp;</p>
  <table width="340" border="1" align="center">
    <tr>
      <td width="81" align="center" scope="col">用户名</td>
      <th width="243" scope="col"><label for="textfield"></label>
      <input type="text" name="textfield" id="textfield" /></th>
    </tr>
    <tr>
      <td align="center">密码</td>
      <td align="center"><label for="textfield2"></label>
      <input type="password" name="textfield2" id="textfield2" /></td>
    </tr>
    <tr>
      <td><input type="submit" name="button" id="button" value="提交" /></td>
      <td><input type="reset" name="button2" id="button2" value="重置" /></td>
    </tr>
  </table>
</form>
<%=(Recordset1.Fields.Item("").Value)%>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
