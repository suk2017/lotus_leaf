<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<% 
    '���ϸ���ҳ��ȡ��Ϣ
    cno=request("cno")     &""
    tno=request("tno")     &""
    maxStuNum=request("maxStuNum")   &""

    '�жϿγ̺��Ƿ����
    set rec = conn.execute("select ID from CoursePlan where ID = "&cno)
    if rec.eof then
        response.redirect "CourseInsert.asp?info=����ʧ�ܣ��γ̺Ų�����"
    end if
    '�жϽ�ʦ���Ƿ����
    set rec2 = conn.execute("select ID from Teacher where ID = "&tno)
    if rec2.eof then
        response.redirect "CourseInsert.asp?info=����ʧ�ܣ���ʦ�Ų�����"
    end if
    '�жϿγ��Ƿ����ſ�
    set rec3 = conn.execute("select ID from Course where ID = "&cno)
    if not rec3.eof then
        response.redirect "CourseInsert.asp?info=����ʧ�ܣ��ÿγ����ſ�"
    end if
    'ִ�в���
    conn.execute("insert into Course values("&cno&", "&tno&", 0, "&maxStuNum&")")
    conn.execute("insert into Course_Class values("&cno&", 10000)")

    '����
    response.Redirect "CourseInsert.asp?info=����ɹ���("&cno&","&tno&",0, "&maxStuNum&")"
    
    rec.close
    set rec=nothing
    rec2.close
    set rec2=nothing
    rec3.close
    set rec3=nothing
    conn.close
    set conn=nothing
%>