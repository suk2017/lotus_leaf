<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<% 
    '���ϸ���ҳ��ȡ��Ϣ
    cname=request("courseName")     &""
    ctype=request("courseType")     &""
    chours=request("courseHours")   &""
    ccredit=request("courseCredit") &""
    cterm=request("courseTerm")     &""

    '�����ݿ� ��ȡ����id
    set rec=conn.execute("select ID from CoursePlan order by ID desc")
    cno=rec(0)
    cno=cno+1

    '�жϿγ������Ƿ�Ϸ�
    set rec2=conn.execute("select * from CourseType where id = "&ctype)
    if rec2.eof then
        response.write "<script>û��������͵Ŀγ�</script>"
    else
        'ִ�в���
        conn.execute("insert into CoursePlan values("&cno&", '"&cname&"', "&ctype&", "&chours&", "&ccredit&", "&cterm&")")
    end if
    
    '����
    response.Redirect "insertCourse.asp?insertSucceed=1&info=("&cno&","&cname&","&ctype&","&chours&","&ccredit&","&cterm&")"
    
    rec.close
    set rec=nothing
    rec2.close
    set rec2=nothing
    conn.close
    set conn=nothing
%>