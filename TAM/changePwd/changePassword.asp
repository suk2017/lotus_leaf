
<!--#include file="include/UserTest.asp"-->
<html>
<head>
    <title>���ݿ����</title>
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
                    &nbsp;&nbsp;ԭ���룺
                    <input type="password" name="oldPassword" />
                    <br />
                    <br />
                    &nbsp;&nbsp;�����룺
                    <input type="password" name="newPassword" />
                    <br />
                    <br />
                    ȷ�����룺
                    <input type="password" name="newPassword2" />
                    <br />
                    <br />
                    <!-- ��֤�� -->
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="submit" value="ȷ��"/>
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
