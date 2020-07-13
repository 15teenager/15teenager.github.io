<!--#include file="config.inc"-->

<%
' Graphical Hit Counter

On Error Resume Next
	
' Declare variables
Dim CounterHits
Dim FixedDigitCount
Dim DigitCount
Dim DigitCountLength
Dim DigitZerosToAdd
Dim DigitZeroCount
Dim ShowDigits

' (FixedDigitCount) will add zeros to the front of your count
' if the count is less then the (FixedDigitCount)
' just like "frontpage" counters let you do
	
FixedDigitCount = 6

'##     打开数据库连接
set conn=server.createobject("adodb.connection")
conn.open ConnString

strSql = "SELECT * FROM Counter"
set rs = conn.Execute (StrSql)

If rs.Eof or rs.Bof then
	CounterHits=0
	strSql = "insert into Counter (count) Values (" & CounterHits & ")"
	conn.Execute (StrSql)	
else
	CounterHits=rs("count")
end if

rs.close
set rs=nothing

CounterHits = CounterHits + 1

strSql = "update Counter set count = " & CounterHits
conn.Execute (StrSql)

'## 关闭数据库连接
conn.close
set conn=nothing

DigitCountLength = Len(CounterHits)

If DigitCountLength < FixedDigitCount Then
	DigitZerosToAdd= FixedDigitCount - DigitCountLength
	DigitZeroCount = 1
	For DigitZeroCount = DigitZeroCount to DigitZerosToAdd
		ShowDigits = ShowDigits & "<img src=""" & DigitPath & "/0.gif"" Alt =""" & CounterHits & " Visitors"" >"
	Next
End If
	
DigitCount = 1
For DigitCount = DigitCount to DigitCountLength
	ShowDigits = ShowDigits & "<img src=""" & DigitPath & "/" & Mid(CounterHits,DigitCount,1) & ".gif"" Alt =""" & CounterHits & " Visitors"">"
Next

Response.Write ShowDigits
%>
