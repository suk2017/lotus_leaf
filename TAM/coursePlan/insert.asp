<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<% 
    '从上个网页获取信息
    cname=request("courseName")     &""
    ctype=request("courseType")     &""
    chours=request("courseHours")   &""
    ccredit=request("courseCredit") &""
    cterm=request("courseTerm")     &""

    '从数据库 获取最新id
    set rec=conn.execute("select ID from CoursePlan order by ID desc")
    cno=rec(0)
    cno=cno+1

    '判断课程类型是否合法
    set rec2=conn.execute("select * from CourseType where id = "&ctype)
    if rec2.eof then
        response.write "<script>没有这个类型的课程</script>"
    else
        '执行插入
        conn.execute("insert into CoursePlan values("&cno&", '"&cname&"', "&ctype&", "&chours&", "&ccredit&", "&cterm&")")
    end if
    
    '返回
    response.Redirect "insertCourse.asp?insertSucceed=1&info=("&cno&","&cname&","&ctype&","&chours&","&ccredit&","&cterm&")"
    
    rec.close
    set rec=nothing
    rec2.close
    set rec2=nothing
    conn.close
    set conn=nothing
%>