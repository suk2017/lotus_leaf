<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>�ſ���Ϣҳ��</title>
    <style type="text/css">
        div {
            margin-right: 10px;
            margin-bottom: 10px;
            float: left
        }

        td {
            text-align: center;
        }

        .table {
            margin-right: 0;
            margin-bottom: 0;
            float: none
        }

        .page{
            margin-left:2%;
        }
    </style>
    <script>
        function myResetBtn(){
            document.getElementById("cno")     .setAttribute("value","");
            document.getElementById("cname")   .setAttribute("value","");
            document.getElementById("tno")     .setAttribute("value","");
            document.getElementById("tname")   .setAttribute("value","");
            document.getElementById("dep0")    .setAttribute("selected","selected");
            document.getElementById("spe0")    .setAttribute("selected","selected");
            document.getElementById("cgrade")   .setAttribute("value","");
        }
    </script>
    <%
        if request("page")&""="" then
            page_cur=1
        else
            page_cur=request("page")
        end if

        recordsPerPage=8
    %>
</head>
<body>
    <%
        '���Ҫɾ��һ����¼
        if request("del")&""="1" then
            sql="delete from CoursePlan where id='" & request("id")  &  "'"
            conn.execute sql
            response.Write "<script>alert("&request("id")&"' ɾ���ɹ���')</script>"
        end if

        '���Ҫ�ύһ����¼
        'if request("exec_upd")&""="1" then
        if request("exec_upd_Btn")&""="�ύ" then
            'sql_upd_exec=" update Student set name='"&request("usname")&"', gender="&request("usgender")&", specialtyID="&request("usspe")&", classID="&classid&", grade="&request("usgrade")&", birth='"&request("usbirth")&"' where id =" &request("usno")
            sql_upd_exec="update CoursePlan set name='"&request("ucname")&"', typeID="&request("uctype")&", hours="&request("uchours")&", credit="&request("uccredit")&", term="&request("ucterm")&" where id="&request("ucno")
            conn.execute sql_upd_exec
             '���ԭ�� "update CoursePlan set name='a', typeID=1, hours=2, credit=3, term=4 where id=0"
        end if
    %>
    <form>
        <div>
            <lavel>�γ̺ţ�</lavel>
            <input type="text" id="cno" name="cno" value="<%=request("cno") %>" />
        </div>
        <div>
            <lavel>�γ�����</lavel>
            <input type="text" id="cname" name="cname" value="<%=request("cname") %>" />
        </div>
        <div>
            <lavel>��ʦ�ţ�</lavel>
            <input type="text" id="tno" name="tno" value="<%=request("tno") %>" />
        </div>
        <div>
            <lavel>��ʦ����</lavel>
            <input type="text" id="tname" name="tname" value="<%=request("tname") %>" />
        </div>
        <div>
            <lavel>ϵ����</lavel>
            <select name="dep">
                <option id="dep0" value=""      <%if request("dep")=""      then %>selected="selected" <%end if %>></option>
                <option id="dep1" value="10001" <%if request("dep")="10001" then %>selected="selected" <%end if %>>�����ϵ</option>
                <option id="dep2" value="10002" <%if request("dep")="10002" then %>selected="selected" <%end if %>>��еϵ</option>
                <option id="dep3" value="10003" <%if request("dep")="10003" then %>selected="selected" <%end if %>>����ϵ</option>
                <option id="dep4" value="10004" <%if request("dep")="10004" then %>selected="selected" <%end if %>>����ϵ�������</option>
            </select>
        </div>
        <div>
            <lavel>רҵ��</lavel>
            <select name="spe">
                <option id="spe0" value=""      <%if request("spe")=""      then %>selected="selected" <%end if %>></option>
                <option id="spe1" value="10001" <%if request("spe")="10001" then %>selected="selected" <%end if %>>�������</option>
                <option id="spe2" value="10002" <%if request("spe")="10002" then %>selected="selected" <%end if %>>����ý���뼼��</option>
                <option id="spe3" value="10003" <%if request("spe")="10003" then %>selected="selected" <%end if %>>�������ѧ�뼼��</option>
                <option id="spe4" value="10004" <%if request("spe")="10004" then %>selected="selected" <%end if %>>��е������켰���Զ���</option>
                <option id="spe5" value="10005" <%if request("spe")="10005" then %>selected="selected" <%end if %>>�Զ���</option>
                <option id="spe6" value="10006" <%if request("spe")="10006" then %>selected="selected" <%end if %>>��ؼ���������</option>
                <option id="spe7" value="10007" <%if request("spe")="10007" then %>selected="selected" <%end if %>>������Ϣ��ѧ�뼼��</option>
                <option id="spe8" value="10008" <%if request("spe")="10008" then %>selected="selected" <%end if %>>ͨ�Ź���</option>
            </select>
        </div>
        <div>
            <lavel>�꼶��</lavel>
            <input type="text" id="cgrade" name="cgrade" value="<%=request("cgrade") %>" />
        </div>

        <div>
            <input type="submit" value="ɸѡ" />
            <input type="reset" value="���" onclick="myResetBtn();" />
        </div>
    </form>
    <hr width="100%" />
    <div class="table">
        <table border="1" width="100%">
            <%
                '׼������
                r1=request("cno")     &""
                r2=request("cname")   &""
                r3=request("tno")     &""
                r4=request("tname")   &""
                r5=request("dep")     &""
                r6=request("spe")     &""
                r7=request("cgrade")  &""
                sqlStr="select distinct c.ID �γ̺�, p.[name] �γ���, c.teacherID ��ʦ��, t.[name] ��ʦ��, c.[count] ѡ������, c.maxStuNum ������, cs.grade �꼶 from Course c, CoursePlan p, Teacher t, Course_Class cc, Class cs, Specialty sy where c.ID=p.ID and c.teacherID = t.ID and c.ID = cc.courseID and cc.classID = cs.ID and cs.specialtyID=sy.ID "

                '��������
                if r1<>"" then
                    sqlStr=sqlStr&" and p.ID = "&r1
                end if
                if r2<>"" then
                    sqlStr=sqlStr&" and p.name like '%"&r2&"%'"
                end if
                if r3<>"" then
                    sqlStr=sqlStr&" and t.ID = "&r3
                end if
                if r4<>"" then
                    sqlStr=sqlStr&" and t.name like '%"&r4&"%'"
                end if
                if r5<>"" then
                    sqlStr=sqlStr&" and sy.deptID = "&r5
                end if
                if r6<>"" then
                    sqlStr=sqlStr&" and cs.specialtyID = "&r6
                end if
                if r7<>"" then
                    sqlStr=sqlStr&" and grade = "&r7
                end if

                'response.write sqlStr

                'ִ�����
                dim rec
                if request("sql")&""<>"" then '��Ҫ��ʹ���ϴε�������ѯ
                    'set rec=conn.execute(request("sql"))
                    sqlllll=request("sql")
                    set rec=server.createobject("adodb.recordset")
                    rec.open sqlllll,conn,3
                    rec.Move (page_cur-1)*recordsPerPage
                else'����ʹ�õ�ǰ������ѯ
                    'set rec=Conn.execute(sqlStr)
                    set rec=server.createobject("adodb.recordset")
                    rec.open sqlStr,conn,3
                    rec.Move (page_cur-1)*recordsPerPage
                end if
            %>
            <tr height="40px">
                <%
                'ѭ�������ͷ 
                for i=0 to 6 'rec.fields.count-1
                    response.write "<th>"&rec.fields(i).name&"</th>"
                next
                response.write "<th colspan='2' align='center'><a href='CourseInsert.asp'>����ſ�</a></th>"
                %>
            </tr>

            <%
                '======ѭ���������=====
                ii=0
                do while (not rec.eof) and (ii < recordsPerPage)
                    'ÿ�м�¼�Ŀ�ʼ
                    response.write "<tr height=""40px"">"
                
                        for i=0 to 6 'rec.fields.count-1
	                        response.write "<td>"&rec(i)&"</td>"
	                    next
                        if request("upd")&""<>"1" then
                            'response.write "<td><a href='CourseEdit.asp?id="&rec(0)&"'>�޸�</a></td>"
                            'response.write "<td><a href='CourseDelete.asp?id="&rec(0)&"'>ɾ��</a></td>"
                            session("CUID")=rec(0)
                            response.write "<td><input type='button' onclick='window.open(""CourseUpdate.asp?id="&rec(0)&""");' value='�޸�' /></td>"
                            response.write "<td><input type='button' onclick='window.open(""CourseDelete.asp?id="&rec(0)&""");' value='ɾ��' /></td>"
                        end if

	                response.write "</tr>"
	                rec.movenext
                    ii=ii+1
	            loop
            %>
        </table>
        <label>��ǰҳ����</label>
        <%
            count=rec.recordcount-1
            pages=count/recordsPerPage + 1
            for i=1 to pages
                if i&""=page_cur&"" then
            %>
        <label class="page">(<%=i %>)</label>
        <%
                else
        %>
        <a class="page" href="Courses.asp?page=<%=i %>"><%=i %></a>
        <%
                end if
            next    
        %>
    </div>

    <%
        '�ͷ���Դ
        rec.close
        set rec=nothing
        Conn.close
        set Conn=nothing
    %>
</body>
