<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<% 
    '从上个网页获取信息
    cno=request("cno")     &""
    tno=request("tno")     &""
    maxStuNum=request("maxStuNum")   &""

    '判断课程号是否存在
    set rec = conn.execute("select ID from CoursePlan where ID = "&cno)
    if rec.eof then
        response.redirect "CourseInsert.asp?info=插入失败！课程号不存在"
    end if
    '判断教师号是否存在
    set rec2 = conn.execute("select ID from Teacher where ID = "&tno)
    if rec2.eof then
        response.redirect "CourseInsert.asp?info=插入失败！教师号不存在"
    end if
    '判断课程是否已排课
    set rec3 = conn.execute("select ID from Course where ID = "&cno)
    if not rec3.eof then
        response.redirect "CourseInsert.asp?info=插入失败！该课程已排课"
    end if
    '执行插入
    conn.execute("insert into Course values("&cno&", "&tno&", 0, "&maxStuNum&")")
    conn.execute("insert into Course_Class values("&cno&", 10000)")

    '返回
    response.Redirect "CourseInsert.asp?info=插入成功！("&cno&","&tno&",0, "&maxStuNum&")"
    
    rec.close
    set rec=nothing
    rec2.close
    set rec2=nothing
    rec3.close
    set rec3=nothing
    conn.close
    set conn=nothing
%>