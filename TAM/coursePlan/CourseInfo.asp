<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>课程信息页面</title>
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
        '如果要删除一条记录
        if request("del")&""="1" then
            sql="delete from CoursePlan where id='" & request("id")  &  "'"
            conn.execute sql
            response.Write "<script>alert("&request("id")&"' 删除成功！')</script>"
        end if

        '如果要提交一条记录
        'if request("exec_upd")&""="1" then
        if request("exec_upd_Btn")&""="提交" then
            'sql_upd_exec=" update Student set name='"&request("usname")&"', gender="&request("usgender")&", specialtyID="&request("usspe")&", classID="&classid&", grade="&request("usgrade")&", birth='"&request("usbirth")&"' where id =" &request("usno")
            sql_upd_exec="update CoursePlan set name='"&request("ucname")&"', typeID="&request("uctype")&", hours="&request("uchours")&", credit="&request("uccredit")&", term="&request("ucterm")&" where id="&request("ucno")
            conn.execute sql_upd_exec
             '语句原型 "update CoursePlan set name='a', typeID=1, hours=2, credit=3, term=4 where id=0"
        end if
    %>
    <form>
        <div>
            <label>课程号：</label><input type="text" size="7" name="cno" id="cno" value="<%=request("cno") %>" />
        </div>
        <div>
            <label>课程名：</label><input type="text" size="7" name="cname" id="cname" value="<%=request("cname") %>" />
        </div>
        <div>
            <label>课程类型</label>
            <select name="ctype" id="ctype">
                <option id="ct" value="" <%if request("ctype")=""  then %>selected="selected" <%end if %>></option>
                <option id="ct1" value="1" <%if request("ctype")="1" then %>selected="selected" <%end if %>>必修</option>
                <option id="ct2" value="2" <%if request("ctype")="2" then %>selected="selected" <%end if %>>选修</option>
                <option id="ct3" value="3" <%if request("ctype")="3" then %>selected="selected" <%end if %>>通识选修</option>
                <option id="ct4" value="4" <%if request("ctype")="4" then %>selected="selected" <%end if %>>通识核心</option>
            </select>
        </div>
        <div>
            <label>学时：</label>
            <input type="number" name="chours" id="chours" value="<%=request("chours") %>" />
            <!--<select name="sspe" id="sspe">
                <option id="ss" value="" <%if request("sspe")=""      then %>selected="selected" <%end if %>></option>
                <option id="ss1" value="10001" <%if request("sspe")="10001" then %>selected="selected" <%end if %>>软件工程</option>
                <option id="ss2" value="10002" <%if request("sspe")="10002" then %>selected="selected" <%end if %>>数字媒体与技术</option>
                <option id="ss3" value="10003" <%if request("sspe")="10003" then %>selected="selected" <%end if %>>计算机科学与技术</option>
                <option id="ss4" value="10004" <%if request("sspe")="10004" then %>selected="selected" <%end if %>>机械设计制造及其自动化</option>
                <option id="ss5" value="10005" <%if request("sspe")="10005" then %>selected="selected" <%end if %>>自动化</option>
                <option id="ss6" value="10006" <%if request("sspe")="10006" then %>selected="selected" <%end if %>>测控技术与仪器</option>
                <option id="ss7" value="10007" <%if request("sspe")="10007" then %>selected="selected" <%end if %>>电子信息科学与技术</option>
                <option id="ss8" value="10008" <%if request("sspe")="10008" then %>selected="selected" <%end if %>>通信工程</option>
            </select>-->
        </div>
        <div>
            <label>学分：</label>
            <input type="number" name="ccredit" id="ccredit" value="<%=request("ccredit") %>" />
            <!--<select name="sclass" id="sclass">
                <option id="sc" value="" <%if request("sclass")=""  then %>selected="selected" <%end if %>></option>
                <option id="sc1" value="1" <%if request("sclass")="1" then %>selected="selected" <%end if %>>1</option>
                <option id="sc2" value="2" <%if request("sclass")="2" then %>selected="selected" <%end if %>>2</option>
                <option id="sc3" value="3" <%if request("sclass")="3" then %>selected="selected" <%end if %>>3</option>
            </select>-->
        </div>
        <div>
            <label>开课学期：</label>
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
            <input type="submit" value="筛选" />
            <input type="reset" value="清空" onclick="myResetBtn();" />
        </div>
    </form>
    <hr width="100%" />
    <div class="table">
        <table border="1" width="107%">
            <%
                '准备数据
                r1=request("cno")        &""
                r2=request("cname")    &""
                r3=request("ctype")&""
                r4=request("chours")      &""
                r5=request("ccredit")  &""
                r6=request("cterm")  &""
                'sqlStr="select Student.ID 学号,Student.name 姓名,Student.gender 性别,Specialty.name 专业, Class.name 班级,Student.grade 年级,Student.birth 出生年月, Class.ID, Specialty.ID from Student,Specialty,Class where Student.specialtyID=Specialty.ID and Student.classID=Class.ID"
                sqlStr="select p.ID 课程号, p.name 课程名, t.name 课程类型, p.hours 学时, p.credit 学分, p.term 开课学期 from CoursePlan p, CourseType t where p.typeID=t.ID"

                '附加条件
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

                '执行语句
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
                '循环输出题头 
                for i=0 to 5 'rec.fields.count-1
                    response.write "<th>"&rec.fields(i).name&"</th>"
                next
                response.write "<th colspan='2' align='center'><a href='insertCourse.asp'>添加课程</a></th>"
                %>
            </tr>

            <%
                '循环输出数据

                '判断是否需要更新
                if request("upd")&""="1" then
                    id_forUpd=request("count")
                else
                    id_forUpd=-1
                end if

                '设置一个步进变量（可以用for循环代替）
                ii=0
                do while (not rec.eof) and (ii < recordsPerPage)
                    '每行记录的开始
                    response.write "<tr height=""40px"">"
                
                    '若当前行不是需要修改的行 
                    if id_forUpd&""<>ii&"" then
                        for i=0 to 5 'rec.fields.count-3
	                        response.write "<td>"&rec(i)&"</td>"
	                    next
                        if request("upd")&""<>"1" then
                            response.write "<td><a href='CourseInfo.asp?upd=1&count="&ii&"&sql="&sqlStr&"&page="&page_cur&"'>修改</a></td>"
                            response.write "<td><a href='CourseInfo.asp?del=1&id="&rec(0)&"&page="&page_cur&"'>删除</a></td>"
                        end if
                    else'若是需要修改的行
                        
                        '输出表格
                        response.write "<form name='update' >"
                        response.write "<input type='hidden' name='ucno' value='"&rec(0)&"' />"
                        response.write "<td>"&rec(0)&"</td>"
                        response.write "<td><input type='text' size='7' name='ucname' value='"&rec(1)&"' /></td>"
                        response.write "<td><select name='uctype'>"
                            response.write "<option value='1' " : if rec(2)="1" then : response.write "selected='selected'" : end if : response.write ">必修</option>"
                            response.write "<option value='2' " : if rec(2)="2" then : response.write "selected='selected'" : end if : response.write ">选修</option>"
                            response.write "<option value='3' " : if rec(2)="3" then : response.write "selected='selected'" : end if : response.write ">通识选修</option>"
                            response.write "<option value='4' " : if rec(2)="4" then : response.write "selected='selected'" : end if : response.write ">通识核心</option>"
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

                        response.write "<td><input type='button' value='取消' onclick=""window.location.href='CourseInfo.asp?sql="&request("sql")&"';return false;""/></td>"
                        response.write "<td><input type='submit' name='exec_upd_Btn' value='提交' /></td>"
                    
                        response.write "</form>"
                    end if

	                response.write "</tr>"
	                rec.movenext
                    ii=ii+1
	            loop
            %>
        </table>
        <label>当前页数：</label>
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
        '释放资源
        rec.close
        set rec=nothing
        Conn.close
        set Conn=nothing
    %>
</body>
