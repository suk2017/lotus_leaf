
<!--#include file="include/UserTest.asp"-->
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <title>����</title>
</head>
    <%
        sname=request("stuName")    &""
        sgender=request("stuGender")&""
        sspe=request("stuSpe")      &""
        sclass=request("stuClass")  &""
        sgrade=request("stuGrade")  &""
        sbirth=request("stuBirth")  &""
        if sname="" then
            sname="����"
        end if
        if sgender="" then
            sgender="0"
        end if
        if sspe="" then
            sspe="1"
        end if
        if sclass="" then
            sclass="1"
        end if
        if sgrade="" then
            sgrade=year(now())
        end if
        if sbirth="" then
            sbirth=(year(now())-18) & "-01-01"
        end if
        
        
    %>
<body>
    <form action="insert.asp" method="get">
        ������<input type="text" name="stuName" value="<%=sname %>"" />
        <br />
        <br />
        �Ա�
        <input type="radio" name="stuGender" value="0" checked/> Ů
        <input type="radio" name="stuGender" value="1"/> ��
        <br />
        <br />
        רҵ��
    <select name="stuSpe">
        <option value="10001">�������  </option>
        <option value="10002">����ý���뼼�� </option>
        <option value="10003">�������ѧ�뼼��   </option>
        <option value="10004">��е������켰���Զ���</option>
        <option value="10005">�Զ���   </option>
        <option value="10006">��ؼ���������   </option>
        <option value="10007">������Ϣ��ѧ�뼼��     </option>
        <option value="10008">ͨ�Ź���    </option>
    </select>
        <br />
        <br />
        �༶��<input type="number" value="1" name="stuClass"/>
        <br />
        <br />
        �꼶��<input type="number" value="<%=sgrade %>" name="stuGrade"/>
        <br />
        <br />
        ���գ�<input type="date" value="<%=sbirth%>" name="stuBirth"/>
        <br />
        <br />
        &nbsp;
        <input type="submit" value="�ύ" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="reset" value="���" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="button" value="����" onclick="window.location.href = 'StudentInfo.asp'; return false;" />
    </form>
    <%
        response.Write(request("info"))
'stu_id=Request("id")

'sql="delete from Student where id='" & stu_id  &  "'"
'response.write sql 
'conn.execute  sql
'Conn.Close
'Set Conn=Nothing
'response.Redirect "StudentInfo.asp"
    %>
</body>
</html>
