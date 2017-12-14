<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>更新排课信息界面</title>
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
    <script>
        //打开连接
        function Conn() { 
            var conn = new ActiveXObject("ADODB.Connection");
            conn.Open("Provider=SQLOLEDB.1; Data Source=127.0.0.1; User ID=user1; Password=123456; Initial Catalog=ATADB");
            return conn;
        } 
        //关闭连接
        function CloseConn() { 
            if(conn != null) { 
                conn.close(); 
                conn = null; 
            } 
        }
        //清空所有筛选条件
        function myResetBtn(){
            document.getElementById("cno")     .setAttribute("value","");
            document.getElementById("cname")   .setAttribute("value","");
            document.getElementById("tno")     .setAttribute("value","");
            document.getElementById("tname")   .setAttribute("value","");
            document.getElementById("dep0")    .setAttribute("selected","selected");
            document.getElementById("spe0")    .setAttribute("selected","selected");
            document.getElementById("cgrade")  .setAttribute("value","");
        }
        //联动更新课程名字
        function selectCName(){
            alert("no");
            Conn();

            

            var cno = document.getElementById("cno").value;

            try { 
                var rs = new ActiveXObject("ADODB.Recordset"); 
                rs.open("select name from CoursePlan where ID = " + cno, conn); 
                document.getElementById("cname").innerText = rs.Fields("name")+"√"
            } catch (e) { 
                document.getElementById("tname").innerText = "X"
                document.write(e.description); 
            } finally { 
                rs.close();
                re=null;
                CloseConn(); 
            }
        }
        //联动更新教师名字
        function selectTName(){
            Conn();
            
            var tno = document.getElementById("tno").value;

            try { 
                var rs = new ActiveXObject("ADODB.Recordset"); 
                rs.open("select name from Teacher where ID = " + tno, conn); 
                document.getElementById("tname").innerText = rs.Fields("name")+"√"
            } catch (e) { 
                document.getElementById("tname").innerText = "X"
                document.write(e.description); 
            } finally { 
                rs.close();
                rs=null;
                CloseConn(); 
            }
        }

    </script>
</head>
<body>
    <%
        if request("succeed") then
            response.write "<script>alert('更新成功！')</script>"
        end if
    %>
    <%
        set rec = conn.execute("select ID,TeacherID, maxStuNum from Course where ID = "&session("CUID"))
        teacherNo=rec(1)
        maxStuNum=rec(2)
    %>
    <form name="u1" action="Update.asp">
        <div>
            <label>课程号：</label>
            <input type="hidden" name="cno" value="<%=session("CUID") %>" />
            <label><%=session("CUID") %></label>
        </div>
        <hr width="100%"/>
        <div>
            <label>教师号：</label>
            <input type="text" name="tno" value="<%=teacherNo %>" />
        </div>
        <div>
            <label>课容量：</label>
            <input type="text" name="maxStuNum" value="<%=maxStuNum %>" />
        </div>
        <div>
            <label>确定修改吗？</label>
            <input type="submit" name="Class" value="确定" />
            &nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="返回" onclick="window.close();" />
        </div>
    </form>
    <hr width="100%"/>

    <form name="u2" action="Update.asp">
        <div>
            <% 
                set rec2 = conn.execute("select ID,name from Class where ID not in (select ClassID from Course_Class where CourseID = "&session("CUID")&")")
            %>
            <input type="hidden" name="cno" value="<%=session("CUID") %>" />
            <label>添加班级：</label>
            <select name="classNo">
            <% do while NOT rec2.eof %>
                    <option value="<%=rec2(0) %>"><%=rec2(1) %></option>
            <% rec2.movenext
               loop %>
            </select>
            <input type="submit" name="Class" value="添加这个班级" />
        </div>
    </form>
    
    <form name="u3" action="Update.asp">
        <div>
            <% 
                set rec3 = conn.execute("select c.ID,c.name from Class c, Course_Class cc where c.ID = cc.ClassID and cc.CourseID = "&session("CUID"))
            %>
            <input type="hidden" name="cno" value="<%=session("CUID") %>" />
            <label>减去班级：</label>
            <select name="classNo">
            <% do while NOT rec3.eof %>
                <option value="<%=rec3(0) %>"><%=rec3(1) %></option>
            <% rec3.movenext
               loop %>
            </select>
            <input type="submit" name="Class" value="减去这个班级" />
        </div>
    </form>


    <%
        '释放资源
        rec.close
        set rec=nothing
        rec2.close
        set rec2=nothing
        rec3.close
        set rec3=nothing
        
        Conn.close
        set Conn=nothing
    %>
</body>
