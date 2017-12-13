<!--#include file="include/connect.asp"-->
<head>
    <title>删除页面</title>
    
</head>
<body>
    <%
        '如果要删除一条记录
        conn.execute "delete from Course_Student where courseID = " & request("id") 
        conn.execute "delete from Course_Class where courseID = " & request("id")
        conn.execute "delete from Course where ID = " & request("id")

        response.write "删除成功!五秒后自动关闭页面"
        response.write "<script>setTimeout(function(){window.close();}, 3000);</script>"
       
        '释放资源
        Conn.close
        set Conn=nothing
    %>
</body>
