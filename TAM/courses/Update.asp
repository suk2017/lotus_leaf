<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>更新界面</title>
    <meta charset="gb2312" />
</head>
<%
    if request("Class")&"" = "确定" then
        conn.execute("update Course set TeacherID = " & request("tno") & ", maxStuNum = " & request("maxStuNum") & " where ID = " & request("cno"))
    elseif request("Class")&"" = "添加这个班级" then
        conn.execute("insert into Course_Class values("&request("cno")&", "&request("classNo")&")")
    elseif request("Class")&"" = "减去这个班级" then
        conn.execute("delete from Course_Class where CourseID = " & request("cno") & " and ClassID = " & request("classNo"))
    end if
    response.redirect "CourseUpdate.asp?succeed=1"

    conn.close
    set conn=nothing
%>