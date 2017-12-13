<!--#include file="include/connect.asp"-->
<head>
    <title>更新界面</title>
    <meta charset="gb2312" />
</head>
<%
    if request("BaseInfo")&"" <> "" then
        conn.execute("update Course set TeacherID = " & request("tno") & ", maxStuNum = " & request("maxStuNum") & " where ID = " & request("cno"))
    elseif request("AddClass")&"" <> "" then
        conn.execute("insert into Course_Class values("&request("cno")&", "&request("classNo")&")")
    elseif request("RmvClass")&"" <> "" then
        conn.execute("delete from Course_Class where CourseID = " & request("cno") & " and ClassID = " & request("classN)
    end if
    response.redirect "CourseUpdate.asp?succeed=1"

    conn.close
    set conn=nothing
%>