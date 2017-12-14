<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>添加排课信息界面</title>
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
            <label>课程号：</label>
            <input type="text" name="cno" value="<%=20001 %>" />
        </div>
        <div>
            <label>教师号：</label>
            <input type="text" name="tno" value="<%=10001 %>" />
        </div>
        <div>
            <label>课容量：</label>
            <input type="text" name="maxStuNum" value="<%=60 %>" />
        </div>
        <div>
            <label>确定添加吗？</label>
            <input type="submit" name="baseInfo" value="确定" />
            &nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="返回" onclick="window.location.href='Courses.asp';" />
        </div>
    </form>
    <hr width="100%"/>
    <br />
    <label>&nbsp;&nbsp;&nbsp;&nbsp;注意： 班级请在添加之后进行修改。 </label>
    <hr />
    <%
        response.write request("info")    
    %>

    <%
        Conn.close
        set Conn=nothing
    %>
</body>
