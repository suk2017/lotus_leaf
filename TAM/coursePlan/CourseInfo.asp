<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>�γ���Ϣҳ��</title>
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
            document.getElementById("ct")      .setAttribute("selected","selected");
            document.getElementById("chours")  .setAttribute("selected","1");
            document.getElementById("ccredit") .setAttribute("value","");
            document.getElementById("cte")     .setAttribute("selected","selected");
        }
    </script>
    <%
        if request("page")&""="" then
            page_cur=1
        else
            page_cur=request("page")
        end if

        recordsPerPage=9
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
            <label>�γ̺ţ�</label><input type="text" size="7" name="cno" id="cno" value="<%=request("cno") %>" />
        </div>
        <div>
            <label>�γ�����</label><input type="text" size="7" name="cname" id="cname" value="<%=request("cname") %>" />
        </div>
        <div>
            <label>�γ�����</label>
            <select name="ctype" id="ctype">
                <option id="ct" value="" <%if request("ctype")=""  then %>selected="selected" <%end if %>></option>
                <option id="ct1" value="1" <%if request("ctype")="1" then %>selected="selected" <%end if %>>����</option>
                <option id="ct2" value="2" <%if request("ctype")="2" then %>selected="selected" <%end if %>>ѡ��</option>
                <option id="ct3" value="3" <%if request("ctype")="3" then %>selected="selected" <%end if %>>ͨʶѡ��</option>
                <option id="ct4" value="4" <%if request("ctype")="4" then %>selected="selected" <%end if %>>ͨʶ����</option>
            </select>
        </div>
        <div>
            <label>ѧʱ��</label>
            <input type="number" name="chours" id="chours" value="<%=request("chours") %>" />
            <!--<select name="sspe" id="sspe">
                <option id="ss" value="" <%if request("sspe")=""      then %>selected="selected" <%end if %>></option>
                <option id="ss1" value="10001" <%if request("sspe")="10001" then %>selected="selected" <%end if %>>�������</option>
                <option id="ss2" value="10002" <%if request("sspe")="10002" then %>selected="selected" <%end if %>>����ý���뼼��</option>
                <option id="ss3" value="10003" <%if request("sspe")="10003" then %>selected="selected" <%end if %>>�������ѧ�뼼��</option>
                <option id="ss4" value="10004" <%if request("sspe")="10004" then %>selected="selected" <%end if %>>��е������켰���Զ���</option>
                <option id="ss5" value="10005" <%if request("sspe")="10005" then %>selected="selected" <%end if %>>�Զ���</option>
                <option id="ss6" value="10006" <%if request("sspe")="10006" then %>selected="selected" <%end if %>>��ؼ���������</option>
                <option id="ss7" value="10007" <%if request("sspe")="10007" then %>selected="selected" <%end if %>>������Ϣ��ѧ�뼼��</option>
                <option id="ss8" value="10008" <%if request("sspe")="10008" then %>selected="selected" <%end if %>>ͨ�Ź���</option>
            </select>-->
        </div>
        <div>
            <label>ѧ�֣�</label>
            <input type="number" name="ccredit" id="ccredit" value="<%=request("ccredit") %>" />
            <!--<select name="sclass" id="sclass">
                <option id="sc" value="" <%if request("sclass")=""  then %>selected="selected" <%end if %>></option>
                <option id="sc1" value="1" <%if request("sclass")="1" then %>selected="selected" <%end if %>>1</option>
                <option id="sc2" value="2" <%if request("sclass")="2" then %>selected="selected" <%end if %>>2</option>
                <option id="sc3" value="3" <%if request("sclass")="3" then %>selected="selected" <%end if %>>3</option>
            </select>-->
        </div>
        <div>
            <label>����ѧ�ڣ�</label>
            <select name="cterm" id="cterm">
                <option id="cte" value="" <%if request("cterm")=""   then %>selected="selected" <%end if %>></option>
                <option id="cte1" value="1" <%if request("cterm")="1"  then %>selected="selected" <%end if %>>1</option>
                <option id="cte2" value="2" <%if request("cterm")="2"  then %>selected="selected" <%end if %>>2</option>
                <option id="cte3" value="3" <%if request("cterm")="3"  then %>selected="selected" <%end if %>>3</option>
                <option id="cte4" value="4" <%if request("cterm")="4"  then %>selected="selected" <%end if %>>4</option>
                <option id="cte5" value="5" <%if request("cterm")="5"  then %>selected="selected" <%end if %>>5</option>
                <option id="cte6" value="6" <%if request("cterm")="6"  then %>selected="selected" <%end if %>>6</option>
                <option id="cte7" value="7" <%if request("cterm")="7"  then %>selected="selected" <%end if %>>7</option>
                <option id="cte8" value="8" <%if request("cterm")="8"  then %>selected="selected" <%end if %>>8</option>
                <option id="cte9" value="9" <%if request("cterm")="9"  then %>selected="selected" <%end if %>>9</option>
                <option id="cte10" value="10" <%if request("cterm")="10" then %>selected="selected" <%end if %>>10</option>
                <option id="cte11" value="11" <%if request("cterm")="11" then %>selected="selected" <%end if %>>11</option>
            </select>
        </div>
        <div>
            <input type="submit" value="ɸѡ" />
            <input type="reset" value="���" onclick="myResetBtn();" />
        </div>
    </form>
    <hr width="100%" />
    <div class="table">
        <table border="1" width="107%">
            <%
                '׼������
                r1=request("cno")        &""
                r2=request("cname")    &""
                r3=request("ctype")&""
                r4=request("chours")      &""
                r5=request("ccredit")  &""
                r6=request("cterm")  &""
                'sqlStr="select Student.ID ѧ��,Student.name ����,Student.gender �Ա�,Specialty.name רҵ, Class.name �༶,Student.grade �꼶,Student.birth ��������, Class.ID, Specialty.ID from Student,Specialty,Class where Student.specialtyID=Specialty.ID and Student.classID=Class.ID"
                sqlStr="select p.ID �γ̺�, p.name �γ���, t.name �γ�����, p.hours ѧʱ, p.credit ѧ��, p.term ����ѧ�� from CoursePlan p, CourseType t where p.typeID=t.ID"

                '��������
                if r1<>"" then
                    sqlStr=sqlStr&" and p.ID = "&r1
                end if
                if r2<>"" then
                    sqlStr=sqlStr&" and p.name like '%"&r2&"%'"
                end if
                if r3<>"" then
                    sqlStr=sqlStr&" and p.typeID = "&r3
                end if
                if r4<>"" then
                    sqlStr=sqlStr&" and p.hours = "&r4
                end if
                if r5<>"" then
                    sqlStr=sqlStr&" and p.credit = "&r5
                end if
                if r6<>"" then
                    sqlStr=sqlStr&" and p.term = "&r6
                end if
                
                'response.write sqlStr

                'ִ�����
                    dim rec
                if request("sql")&""<>"" then 
                    'set rec=conn.execute(request("sql"))
                    sqlllll=request("sql")
                    set rec=server.createobject("adodb.recordset")
                    rec.open sqlllll,conn,3
                    rec.Move (page_cur-1)*recordsPerPage
                else
                    'set rec=Conn.execute(sqlStr)
                    set rec=server.createobject("adodb.recordset")
                    rec.open sqlStr,conn,3
                    rec.Move (page_cur-1)*recordsPerPage
                end if
            %>
            <tr height="40px">
                <%
                'ѭ�������ͷ 
                for i=0 to 5 'rec.fields.count-1
                    response.write "<th>"&rec.fields(i).name&"</th>"
                next
                response.write "<th colspan='2' align='center'><a href='insertCourse.asp'>��ӿγ�</a></th>"
                %>
            </tr>

            <%
                'ѭ���������

                '�ж��Ƿ���Ҫ����
                if request("upd")&""="1" then
                    id_forUpd=request("count")
                else
                    id_forUpd=-1
                end if

                '����һ������������������forѭ�����棩
                ii=0
                do while (not rec.eof) and (ii < recordsPerPage)
                    'ÿ�м�¼�Ŀ�ʼ
                    response.write "<tr height=""40px"">"
                
                    '����ǰ�в�����Ҫ�޸ĵ��� 
                    if id_forUpd&""<>ii&"" then
                        for i=0 to 5 'rec.fields.count-3
	                        response.write "<td>"&rec(i)&"</td>"
	                    next
                        if request("upd")&""<>"1" then
                            response.write "<td><a href='CourseInfo.asp?upd=1&count="&ii&"&sql="&sqlStr&"&page="&page_cur&"'>�޸�</a></td>"
                            response.write "<td><a href='CourseInfo.asp?del=1&id="&rec(0)&"&page="&page_cur&"'>ɾ��</a></td>"
                        end if
                    else'������Ҫ�޸ĵ���
                        
                        '������
                        response.write "<form name='update' >"
                        response.write "<input type='hidden' name='ucno' value='"&rec(0)&"' />"
                        response.write "<td>"&rec(0)&"</td>"
                        response.write "<td><input type='text' size='7' name='ucname' value='"&rec(1)&"' /></td>"
                        response.write "<td><select name='uctype'>"
                            response.write "<option value='1' " : if rec(2)="1" then : response.write "selected='selected'" : end if : response.write ">����</option>"
                            response.write "<option value='2' " : if rec(2)="2" then : response.write "selected='selected'" : end if : response.write ">ѡ��</option>"
                            response.write "<option value='3' " : if rec(2)="3" then : response.write "selected='selected'" : end if : response.write ">ͨʶѡ��</option>"
                            response.write "<option value='4' " : if rec(2)="4" then : response.write "selected='selected'" : end if : response.write ">ͨʶ����</option>"
                        response.write "</select></td>"
                        response.write "<td><input type='number' name='uchours' value='"&rec(3)&"' />"
                        response.write "<td><select name='uccredit'>"
                            response.write "<option value='1'  "   : if rec(4)="1"   then : response.write "selected='selected'" : end if : response.write ">1</option>"
                            response.write "<option value='1.5' "  : if rec(4)="1.5" then : response.write "selected='selected'" : end if : response.write ">1.5</option>"
                            response.write "<option value='2'  "   : if rec(4)="2"   then : response.write "selected='selected'" : end if : response.write ">2</option>"
                            response.write "<option value='2.5' "  : if rec(4)="2.5" then : response.write "selected='selected'" : end if : response.write ">2.5</option>"
                            response.write "<option value='3'  "   : if rec(4)="3"   then : response.write "selected='selected'" : end if : response.write ">3</option>"
                            response.write "<option value='3.5' "  : if rec(4)="3.5" then : response.write "selected='selected'" : end if : response.write ">3.5</option>"
                            response.write "<option value='4'  "   : if rec(4)="4"   then : response.write "selected='selected'" : end if : response.write ">4</option>"
                            response.write "<option value='4.5' "  : if rec(4)="4.5" then : response.write "selected='selected'" : end if : response.write ">4.5</option>"
                            response.write "<option value='5'  "   : if rec(4)="5"   then : response.write "selected='selected'" : end if : response.write ">5</option>"
                        response.write "</select></td>"
                        response.write "<td><select name='ucterm'>"
                            response.write "<option value='1'  " : if rec(5)="1"  then : response.write "selected='selected'" : end if : response.write ">1</option>"
                            response.write "<option value='2'  " : if rec(5)="2"  then : response.write "selected='selected'" : end if : response.write ">2</option>"
                            response.write "<option value='3'  " : if rec(5)="3"  then : response.write "selected='selected'" : end if : response.write ">3</option>"
                            response.write "<option value='4'  " : if rec(5)="4"  then : response.write "selected='selected'" : end if : response.write ">4</option>"
                            response.write "<option value='5'  " : if rec(5)="5"  then : response.write "selected='selected'" : end if : response.write ">5</option>"
                            response.write "<option value='6'  " : if rec(5)="6"  then : response.write "selected='selected'" : end if : response.write ">6</option>"
                            response.write "<option value='7'  " : if rec(5)="7"  then : response.write "selected='selected'" : end if : response.write ">7</option>"
                            response.write "<option value='8'  " : if rec(5)="8"  then : response.write "selected='selected'" : end if : response.write ">8</option>"
                            response.write "<option value='9'  " : if rec(5)="9"  then : response.write "selected='selected'" : end if : response.write ">9</option>"
                            response.write "<option value='10' " : if rec(5)="10" then : response.write "selected='selected'" : end if : response.write ">10</option>"
                            response.write "<option value='11' " : if rec(5)="11" then : response.write "selected='selected'" : end if : response.write ">11</option>"
                        response.write "</select></td>"

                        response.write "<td><input type='button' value='ȡ��' onclick=""window.location.href='CourseInfo.asp?sql="&request("sql")&"';return false;""/></td>"
                        response.write "<td><input type='submit' name='exec_upd_Btn' value='�ύ' /></td>"
                    
                        response.write "</form>"
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
        <a class="page" href="CourseInfo.asp?page=<%=i %>"><%=i %></a>
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
