<!--#include file="include/connect.asp"-->
<head>
    <title>ɾ��ҳ��</title>
    
</head>
<body>
    <%
        '���Ҫɾ��һ����¼
        conn.execute "delete from Course_Student where courseID = " & request("id") 
        conn.execute "delete from Course_Class where courseID = " & request("id")
        conn.execute "delete from Course where ID = " & request("id")

        response.write "ɾ���ɹ�!������Զ��ر�ҳ��"
        response.write "<script>setTimeout(function(){window.close();}, 3000);</script>"
       
        '�ͷ���Դ
        Conn.close
        set Conn=nothing
    %>
</body>
