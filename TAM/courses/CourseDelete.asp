<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>ɾ���ſ���Ϣҳ��</title>
    <meta charset="gb2312" />
    <style type="text/css">
        div {
            margin-right: 10px;
            margin-bottom: 10px;
            float: left
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
    <div class="table">
        <table border="1" width="100%">
            <%
                sqlStr="select distinct c.ID �γ̺�, p.[name] �γ���, c.teacherID ��ʦ��, t.[name] ��ʦ��, c.[count] ѡ������, c.maxStuNum ������, cs.grade �꼶 from Course c, CoursePlan p, Teacher t, Course_Class cc, Class cs, Specialty sy where c.ID=p.ID and c.teacherID = t.ID and c.ID = cc.courseID and cc.classID = cs.ID and cs.specialtyID=sy.ID "
                sqlStr=sqlStr&" and c.ID = "&request("id")
                
                'response.write sqlStr

                'ִ�����
                set rec=Conn.execute(sqlStr)
            %>
            <tr height="40px">
                <%
                'ѭ�������ͷ 
                for i=0 to 6 'rec.fields.count-1
                    response.write "<th>"&rec.fields(i).name&"</th>"
                next
                %>
            </tr>

            <%
                '======ѭ���������=====
                response.write "<tr height=""40px"">"
                
                    for i=0 to 6 'rec.fields.count-1
	                    response.write "<td>"&rec(i)&"</td>"
	                next

	            response.write "</tr>"
            %>
        </table>
        <div>
            <label>ȷ��ɾ����</label>
            <form>
                <input type="button" value="ȷ��" onclick="window.location.href='Delete.asp?id=<%=rec(0) %>;'" />
                <input type="button" value="ȡ��" onclick="window.close();" />
            </form>
        </div>
    </div>

    <%
        '�ͷ���Դ
        rec.close
        set rec=nothing
        Conn.close
        set Conn=nothing
    %>
</body>
