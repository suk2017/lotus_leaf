<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>�����ſ���Ϣ����</title>
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
        //������
        function Conn() { 
            var conn = new ActiveXObject("ADODB.Connection");
            conn.Open("Provider=SQLOLEDB.1; Data Source=127.0.0.1; User ID=user1; Password=123456; Initial Catalog=ATADB");
            return conn;
        } 
        //�ر�����
        function CloseConn() { 
            if(conn != null) { 
                conn.close(); 
                conn = null; 
            } 
        }
        //�������ɸѡ����
        function myResetBtn(){
            document.getElementById("cno")     .setAttribute("value","");
            document.getElementById("cname")   .setAttribute("value","");
            document.getElementById("tno")     .setAttribute("value","");
            document.getElementById("tname")   .setAttribute("value","");
            document.getElementById("dep0")    .setAttribute("selected","selected");
            document.getElementById("spe0")    .setAttribute("selected","selected");
            document.getElementById("cgrade")  .setAttribute("value","");
        }
        //�������¿γ�����
        function selectCName(){
            alert("no");
            Conn();

            

            var cno = document.getElementById("cno").value;

            try { 
                var rs = new ActiveXObject("ADODB.Recordset"); 
                rs.open("select name from CoursePlan where ID = " + cno, conn); 
                document.getElementById("cname").innerText = rs.Fields("name")+"��"
            } catch (e) { 
                document.getElementById("tname").innerText = "X"
                document.write(e.description); 
            } finally { 
                rs.close();
                re=null;
                CloseConn(); 
            }
        }
        //�������½�ʦ����
        function selectTName(){
            Conn();
            
            var tno = document.getElementById("tno").value;

            try { 
                var rs = new ActiveXObject("ADODB.Recordset"); 
                rs.open("select name from Teacher where ID = " + tno, conn); 
                document.getElementById("tname").innerText = rs.Fields("name")+"��"
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
            response.write "<script>alert('���³ɹ���')</script>"
        end if
    %>
    <%
        set rec = conn.execute("select ID,TeacherID, maxStuNum from Course where ID = "&session("CUID"))
        teacherNo=rec(1)
        maxStuNum=rec(2)
    %>
    <form name="u1" action="Update.asp">
        <div>
            <label>�γ̺ţ�</label>
            <input type="hidden" name="cno" value="<%=session("CUID") %>" />
            <label><%=session("CUID") %></label>
        </div>
        <hr width="100%"/>
        <div>
            <label>��ʦ�ţ�</label>
            <input type="text" name="tno" value="<%=teacherNo %>" />
        </div>
        <div>
            <label>��������</label>
            <input type="text" name="maxStuNum" value="<%=maxStuNum %>" />
        </div>
        <div>
            <label>ȷ���޸���</label>
            <input type="submit" name="Class" value="ȷ��" />
            &nbsp;&nbsp;&nbsp;&nbsp;
            <input type="button" value="����" onclick="window.close();" />
        </div>
    </form>
    <hr width="100%"/>

    <form name="u2" action="Update.asp">
        <div>
            <% 
                set rec2 = conn.execute("select ID,name from Class where ID not in (select ClassID from Course_Class where CourseID = "&session("CUID")&")")
            %>
            <input type="hidden" name="cno" value="<%=session("CUID") %>" />
            <label>��Ӱ༶��</label>
            <select name="classNo">
            <% do while NOT rec2.eof %>
                    <option value="<%=rec2(0) %>"><%=rec2(1) %></option>
            <% rec2.movenext
               loop %>
            </select>
            <input type="submit" name="Class" value="�������༶" />
        </div>
    </form>
    
    <form name="u3" action="Update.asp">
        <div>
            <% 
                set rec3 = conn.execute("select c.ID,c.name from Class c, Course_Class cc where c.ID = cc.ClassID and cc.CourseID = "&session("CUID"))
            %>
            <input type="hidden" name="cno" value="<%=session("CUID") %>" />
            <label>��ȥ�༶��</label>
            <select name="classNo">
            <% do while NOT rec3.eof %>
                <option value="<%=rec3(0) %>"><%=rec3(1) %></option>
            <% rec3.movenext
               loop %>
            </select>
            <input type="submit" name="Class" value="��ȥ����༶" />
        </div>
    </form>


    <%
        '�ͷ���Դ
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
