<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <title>����</title>
    <style>
        td{
            
        }
    </style>
</head>
<%
        '���ڼ�¼�ϴε�ֵ ���Ǵ��ϸ���ҳ��ȡֵ
        cname=request("courseName")     &""
        ctype=request("courseType")     &""
        chours=request("courseHours")   &""
        ccredit=request("courseCredit") &""
        cterm=request("courseTerm")     &""

        '����������ʹ��Ϸ�
        if cname="" then
            cname="�γ���"
        end if
        if ctype="" then
            ctype="1"
        end if
        if chours="" then
            chours="30"
        end if
        if ccredit="" then
            ccredit="1"
        end if
        if cterm="" then
            cterm="1"
        end if
        
        
        
%>
<body>

    <form action="insert.asp" method="get">
        <table>
            <tr>
                <td>
                    <label>�γ���&nbsp;&nbsp;&nbsp;��</label></td>
                <td>
                    <input type='text' size='7' name='courseName' value='�γ���' /></td>

            </tr>
            <tr>
                <td>
                    <label>�γ����ͣ�</label></td>
                <td>
                    <select name='courseType'>
                        <option value='1'>����</option>
                        <option value='2'>ѡ��</option>
                        <option value='3'>ͨʶѡ��</option>
                        <option value='4'>ͨʶ����</option>
                    </select><td>
            </tr>
            <tr>
                <td>
                    <label>ѧʱ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</label></td>
                <td>
                    <input type='number' name='courseHours' value='30' /></td>
            </tr>
            <tr>
                <td>
                    <label>ѧ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</label></td>
                <td>
                    <select name='courseCredit'>
                        <option value='1'>1</option>
                        <option value='1.5'>1.5</option>
                        <option value='2'>2</option>
                        <option value='2.5'>2.5</option>
                        <option value='3'>3</option>
                        <option value='3.5'>3.5</option>
                        <option value='4'>4</option>
                        <option value='4.5'>4.5</option>
                        <option value='5'>5</option>
                    </select></td>
            </tr>
            <tr>
                <td>
                    <label>ѧ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</label></td>
                <td>
                    <select name='courseTerm'>
                        <option value='1'>1</option>
                        <option value='2'>2</option>
                        <option value='3'>3</option>
                        <option value='4'>4</option>
                        <option value='5'>5</option>
                        <option value='6'>6</option>
                        <option value='7'>7</option>
                        <option value='8'>8</option>
                        <option value='9'>9</option>
                        <option value='10'>10</option>
                        <option value='11'>11</option>
                    </select></td>
        </table>
        </tr>
        <br />
        <br />
        <br />
        &nbsp;
        <input type="submit" value="�ύ" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="reset" value="���" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="button" value="����" onclick="window.location.href = 'CourseInfo.asp'; return false;" />
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
