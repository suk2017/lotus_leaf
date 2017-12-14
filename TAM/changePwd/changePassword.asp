
<!--#include file="include/UserTest.asp"-->
<html>
<head>
    <title>数据库访问</title>
    <meta charset="gb2312" />
</head>
<body>
    <table width="100%">
        <tr>
            <td align="center">
                <br />
                <br />
                <br />
                <br />
                <form action="fun.asp">
                    &nbsp;&nbsp;原密码：
                    <input type="password" name="oldPassword" />
                    <br />
                    <br />
                    &nbsp;&nbsp;新密码：
                    <input type="password" name="newPassword" />
                    <br />
                    <br />
                    确认密码：
                    <input type="password" name="newPassword2" />
                    <br />
                    <br />
                    <!-- 验证码 -->
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="submit" value="确定"/>
                    &nbsp;&nbsp;&nbsp;
                    <input type="reset" />
                </form>
            </td>
        </tr>
    </table>
    <hr />
    <%
        if request("info")&"" <> "" then
            response.write request("info")
        end if
    %>
</body>
</html>
