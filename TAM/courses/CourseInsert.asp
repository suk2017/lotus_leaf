<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>����ſ���Ϣ����</title>
    <meta charset="gb2312" />
    <style type="text/css">
        div {
            margin-right: 10px;
            margin-bottom: 10px;
            float: none
        }

        .table {
            margin-right: 0;
            margin-bottom: 0;
            float: none
        }

        td {
            text-align: center;
        }
    </style>
</head>
<body>
    <form name="u1" action="Insert.asp">
        <div>
            <label>�γ̺ţ�</label>
            <input type="text" name="cno" value="<%=20001 %>" />
        </div>
        <div>
            <label>��ʦ�ţ�</label>
            <input type="text" name="tno" value="<%=10001 %>" />
        </div>
        <div>
            <label>��������</label>
            <input type="text" name="maxStuNum" value="<%=60 %>" />
        </div>
        <div>
            <label>ȷ�������</label>
            <input type="submit" name="baseInfo" value="ȷ��" />
            &nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="����" onclick="window.location.href='Courses.asp';" />
        </div>
    </form>
    <hr width="100%"/>
    <br />
    <label>&nbsp;&nbsp;&nbsp;&nbsp;ע�⣺ �༶�������֮������޸ġ� </label>
    <hr />
    <%
        response.write request("info")    
    %>

    <%
        Conn.close
        set Conn=nothing
    %>
</body>
