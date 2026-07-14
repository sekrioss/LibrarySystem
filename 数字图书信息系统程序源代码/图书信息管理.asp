<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/TuShuManager.asp" -->
<%
Dim Recordset2
Dim Recordset2_cmd
Dim Recordset2_numRows

Set Recordset2_cmd = Server.CreateObject ("ADODB.Command")
Recordset2_cmd.ActiveConnection = MM_TuShuManager_STRING
Recordset2_cmd.CommandText = "SELECT * FROM dbo.BookManger ORDER BY Id ASC" 
Recordset2_cmd.Prepared = true

Set Recordset2 = Recordset2_cmd.Execute
Recordset2_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset2_numRows = Recordset2_numRows + Repeat1__numRows
%>
<%
Dim MM_paramName 
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<style type="text/css">
body,td,th {
	font-family: "宋体";
	font-size: 24px;
}
</style>
</head>

<body>
<table width="1000" border="1" align="center">
  <caption>
    图书信息表
  </caption>
  <tr>
    <th scope="col">图书编码</th>
    <th scope="col">图书书名</th>
    <th scope="col">图书ISBN码</th>
    <th scope="col">图书借阅日期</th>
    <th scope="col">借阅人</th>
    <th scope="col">编辑</th>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset2.EOF)) 
%>
    <tr>
      <td><%=(Recordset2.Fields.Item("Id").Value)%></td>
      <td><%=(Recordset2.Fields.Item("bmmname").Value)%></td>
      <td><%=(Recordset2.Fields.Item("bmisbn").Value)%></td>
      <td><%=(Recordset2.Fields.Item("bmidate").Value)%></td>
      <td><%=(Recordset2.Fields.Item("bmzz").Value)%></td>
      <td><a href="insert.asp?<%= Server.HTMLEncode(MM_keepURL) & MM_joinChar(MM_keepURL) & "Id=" & Recordset2.Fields.Item("Id").Value %>">插入</a></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset2.MoveNext()
Wend
%>
</table>
</body>
</html>
<%
Recordset2.Close()
Set Recordset2 = Nothing
%>