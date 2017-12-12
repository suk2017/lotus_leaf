<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <title>插入</title>
    <style>
        td{
            
        }
    </style>
</head>
<%
        '用于记录上次的值 不是从上个网页获取值
        cname=request("courseName")     &""
        ctype=request("courseType")     &""
        chours=request("courseHours")   &""
        ccredit=request("courseCredit") &""
        cterm=request("courseTerm")     &""

        '处理数据以使其合法
        if cname="" then
            cname="课程名"
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
                    <label>课程名&nbsp;&nbsp;&nbsp;：</label></td>
                <td>
                    <input type='text' size='7' name='courseName' value='课程名' /></td>

            </tr>
            <tr>
                <td>
                    <label>课程类型：</label></td>
                <td>
                    <select name='courseType'>
                        <option value='1'>必修</option>
                        <option value='2'>选修</option>
                        <option value='3'>通识选修</option>
                        <option value='4'>通识核心</option>
                    </select><td>
            </tr>
            <tr>
                <td>
                    <label>学时&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;：</label></td>
                <td>
                    <input type='number' name='courseHours' value='30' /></td>
            </tr>
            <tr>
                <td>
                    <label>学分&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;：</label></td>
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
                    <label>学期&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;：</label></td>
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
        <input type="submit" value="提交" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="reset" value="清空" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="button" value="返回" onclick="window.location.href = 'CourseInfo.asp'; return false;" />
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
