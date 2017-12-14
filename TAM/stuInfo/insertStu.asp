
<!--#include file="include/UserTest.asp"-->
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <title>插入</title>
</head>
    <%
        sname=request("stuName")    &""
        sgender=request("stuGender")&""
        sspe=request("stuSpe")      &""
        sclass=request("stuClass")  &""
        sgrade=request("stuGrade")  &""
        sbirth=request("stuBirth")  &""
        if sname="" then
            sname="姓名"
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
        姓名：<input type="text" name="stuName" value="<%=sname %>"" />
        <br />
        <br />
        性别：
        <input type="radio" name="stuGender" value="0" checked/> 女
        <input type="radio" name="stuGender" value="1"/> 男
        <br />
        <br />
        专业：
    <select name="stuSpe">
        <option value="10001">软件工程  </option>
        <option value="10002">数字媒体与技术 </option>
        <option value="10003">计算机科学与技术   </option>
        <option value="10004">机械设计制造及其自动化</option>
        <option value="10005">自动化   </option>
        <option value="10006">测控技术与仪器   </option>
        <option value="10007">电子信息科学与技术     </option>
        <option value="10008">通信工程    </option>
    </select>
        <br />
        <br />
        班级：<input type="number" value="1" name="stuClass"/>
        <br />
        <br />
        年级：<input type="number" value="<%=sgrade %>" name="stuGrade"/>
        <br />
        <br />
        生日：<input type="date" value="<%=sbirth%>" name="stuBirth"/>
        <br />
        <br />
        &nbsp;
        <input type="submit" value="提交" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="reset" value="清空" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        <input type="button" value="返回" onclick="window.location.href = 'StudentInfo.asp'; return false;" />
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
