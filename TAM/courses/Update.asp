<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>���½���</title>
    <meta charset="gb2312" />
</head>
<%
    if request("Class")&"" = "ȷ��" then
        conn.execute("update Course set TeacherID = " & request("tno") & ", maxStuNum = " & request("maxStuNum") & " where ID = " & request("cno"))
    elseif request("Class")&"" = "�������༶" then
        conn.execute("insert into Course_Class values("&request("cno")&", "&request("classNo")&")")
    elseif request("Class")&"" = "��ȥ����༶" then
        conn.execute("delete from Course_Class where CourseID = " & request("cno") & " and ClassID = " & request("classNo"))
    end if
    response.redirect "CourseUpdate.asp?succeed=1"

    conn.close
    set conn=nothing
%>