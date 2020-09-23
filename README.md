<div align="center">

## Simplest Multiple \- Column sorting \(Up & Down\) in ASP


</div>

### Description

This is the simplest/fastest code that shows how to implement multiple-column sorting (Up & Down) in ASP.

Excellent for ASP Database Begineers
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sine45](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sine45.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\), HTML, VbScript \(browser/client side\)

**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sine45-simplest-multiple-column-sorting-up-down-in-asp__4-8405/archive/master.zip)





### Source Code

```
<style>
th{font-family:arial;font-size:10pt;}
td{font-family:verdana;font-size:9pt;}
</style>
<%
dim conn, connString
connString = "nwind"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connstring
dim rs, sql, url,fname,sort,lastsort,thissort
url = Request.ServerVariables("URL")
sql = "SELECT CompanyName, ContactName, Address, City, Phone From Customers"
sort = lcase(request("sort"))
lastsort = lcase(request("lastsort"))
if sort<>"" then
	if lastsort=sort then
		thissort = sort & " desc"
	elseif instr(lastsort,sort & " desc") then
		thissort = replace(lastsort,sort & " desc",sort)
	elseif instr(lastsort,sort) then
		thissort = replace(lastsort,sort,sort & " desc")
	elseif lastsort<>"" then
		thissort = lastsort & "," & sort
	else
		thissort = sort
	end if
	sql = sql & " ORDER BY " & thissort
end if
Response.Write "<p><b><font color=blue>ORDER BY</font>:</b> " & thissort & "</p>"
Response.Write "&lt;a href=""" & url & """>Reset Order</a>"
'Response.End
set rs = conn.Execute(SQL)
'print headers
Response.Write "<table border=1><tr>"
for i=0 to rs.fields.count - 1
	fname = rs.fields(i).name
	Response.Write "<th>&lt;a href=""" & url & "?sort="& fname &"&lastsort="& thissort & """>" & fname
	if instr(thissort,lcase(fname & " desc")) then
		Response.Write " -"
	elseif instr(thissort,lcase(fname)) then
		Response.Write " +"
	end if
	Response.Write "</th>"
next
Response.Write "</tr>"
'print recs
do while not rs.eof
	Response.Write "<tr>"
	for i=0 to rs.fields.count - 1
		Response.Write "<td>" & rs(i) & "</td>"
	next
	Response.Write "</tr>"
	rs.movenext
loop
Response.Write "</table>"
rs.close
conn.Close
set rs = nothing
set conn = nothing
%>
```

