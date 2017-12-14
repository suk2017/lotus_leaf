<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<% 
    '获取信息
    sname=request("stuName")    &""
    sgender=request("stuGender")&""
    sspe=request("stuSpe")      &""
    sclass=request("stuClass")  &""
    sgrade=request("stuGrade")  &""
    sbirth=request("stuBirth")  &""
    sno=""

    '计算sno前六位
    if len(sgrade)=4 then
        sno = sgrade
    end if
    spe=sspe-10000
    if len(spe)=2 then
        sno = sno & spe
    else
        if len(spe)=1 then
            sno = sno & "0" & spe
        end if
    end if
    
    '计算sno后三位
    set rec=Conn.execute("select id from Student where id >= " & sno & "000 and id < " & (sno+1) & "000 order by id desc")
    if rec.eof then
        sno=sno & "000"
    else
        sno=rec(0)+1 & ""
    end if
    
    '计算班级号
    set rec=Conn.execute("select id from Class where grade = "&sgrade&" and specialtyID = "&sspe&" and number = "&sclass)
    if rec.eof then
        succeed=0
        str="没有这个班级 添加失败"
    else
        sclass=rec(0)
        str= sno&"_"&sname&"_"&sgender&"_"&sspe&"_"&sclass&"_"&sgrade&"_"&sbirth&" 添加成功"
        conn.execute("insert into Student values("&sno&", '"&sname&"', "&sgender&", "&sspe&", "&sclass&", "&sgrade&", 123, '"&sbirth&"')")
    end if
    
    '返回
    response.Redirect "insertStu.asp?info="&str
    
    conn.close
    set conn=nothing
%>