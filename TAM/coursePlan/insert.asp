<!--#include file="../include/connect.asp"-->
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
    if rec2.eof
        response.write "<script>û��������͵Ŀγ�</script>"
    else
        'ִ�в���
        conn.execute("insert into CoursePlan values("&cno&", '"&cname&"', "&ctype&", "&chours&", "&ccredit&", "&cterm&")")
    end if
    
    
    '����snoǰ��λ
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
    
    '����sno����λ
    set rec=Conn.execute("select id from Student where id >= " & sno & "000 and id < " & (sno+1) & "000 order by id desc")
    if rec.eof then
        sno=sno & "000"
    else
        sno=rec(0)+1 & ""
    end if
    
    '����༶��
    set rec=Conn.execute("select id from Class where grade = "&sgrade&" and specialtyID = "&sspe&" and number = "&sclass)
    if rec.eof then
        succeed=0
        str="û������༶ ���ʧ��"
    else
        sclass=rec(0)
        str= sno&"_"&sname&"_"&sgender&"_"&sspe&"_"&sclass&"_"&sgrade&"_"&sbirth&" ��ӳɹ�"
        conn.execute("insert into Student values("&sno&", '"&sname&"', "&sgender&", "&sspe&", "&sclass&", "&sgrade&", 123, '"&sbirth&"')")
    end if
    
    '����
    response.Redirect "insertStu.asp?info="&str
    
    conn.close
    set conn=nothing
%>