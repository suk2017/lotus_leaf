<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>删除排课信息页面</title>
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
                sqlStr="select distinct c.ID 课程号, p.[name] 课程名, c.teacherID 教师号, t.[name] 教师名, c.[count] 选课人数, c.maxStuNum 课容量, cs.grade 年级 from Course c, CoursePlan p, Teacher t, Course_Class cc, Class cs, Specialty sy where c.ID=p.ID and c.teacherID = t.ID and c.ID = cc.courseID and cc.classID = cs.ID and cs.specialtyID=sy.ID "
                sqlStr=sqlStr&" and c.ID = "&request("id")
                
                'response.write sqlStr

                '执行语句
                set rec=Conn.execute(sqlStr)
            %>
            <tr height="40px">
                <%
                '循环输出题头 
                for i=0 to 6 'rec.fields.count-1
                    response.write "<th>"&rec.fields(i).name&"</th>"
                next
                %>
            </tr>

            <%
                '======循环输出数据=====
                response.write "<tr height=""40px"">"
                
                    for i=0 to 6 'rec.fields.count-1
	                    response.write "<td>"&rec(i)&"</td>"
	                next

	            response.write "</tr>"
            %>
        </table>
        <div>
            <label>确定删除吗？</label>
            <form>
                <input type="button" value="确定" onclick="window.location.href='Delete.asp?id=<%=rec(0) %>;'" />
                <input type="button" value="取消" onclick="window.close();" />
            </form>
        </div>
    </div>

    <%
        '释放资源
        rec.close
        set rec=nothing
        Conn.close
        set Conn=nothing
    %>
</body>
