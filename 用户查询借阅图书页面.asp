<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/yyy.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_yyy_STRING
Recordset1_cmd.CommandText = "SELECT * FROM dbo.books" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
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
</head>

<body>
<table width="549" border="1" align="center">
  <tr>
    <th width="539" height="79" bgcolor="#999999" scope="col">图书信息查询表</th>
  </tr>
</table>
<p>&nbsp;</p>
<table width="1290" border="1" align="center">
  <tr>
    <th width="210" scope="col">图书ISBN</th>
    <th width="210" scope="col">书名</th>
    <th width="220" scope="col">图书价格</th>
    <th width="260" scope="col">图书借阅人</th>
    <th width="210" scope="col">图书借阅日期</th>
    <th width="140" scope="col">借阅图书</th>
  </tr>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
  <tr>
    <td><%=(Recordset1.Fields.Item("bookISBN").Value)%></td>
    <td><%=(Recordset1.Fields.Item("bookNAME").Value)%></td>
    <td><%=(Recordset1.Fields.Item("bookPRICE").Value)%></td>
    <td><%=(Recordset1.Fields.Item("bookjieyueren").Value)%></td>
    <td><%=(Recordset1.Fields.Item("bookdata").Value)%></td>
    <td align="center"><a href="图书信息插入页面.asp?<%= Server.HTMLEncode(MM_keepNone) & MM_joinChar(MM_keepNone) & "bookISBN=" & Recordset1.Fields.Item("bookISBN").Value %>"> 登记</a></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
