<!--#include file="include/connect.asp"-->
<head>
    <title>课程信息页面</title>
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
    <script>
        function myResetBtn(){
            document.getElementById("sno")     .setAttribute("value","");
            document.getElementById("sname")   .setAttribute("value","");
            document.getElementById("sg") .setAttribute("selected","selected");
            document.getElementById("ss") .setAttribute("selected","selected");
            document.getElementById("sc") .setAttribute("selected","selected");
            document.getElementById("sgrade")  .setAttribute("value","");
            document.getElementById("sbirth")  .setAttribute("value","");
        }
    </script>
</head>
<body>
    <%
        '如果要删除一条记录
        if request("del")&""="1" then
            sql="delete from Student where id='" & request("id")  &  "'"
            conn.execute sql
            response.Write "<script>alert("&request("id")&"' 删除成功！')</script>"
        end if

        '如果要提交一条记录
        'if request("exec_upd")&""="1" then
        if request("exec_upd_Btn")&""="提交" then
            dim rec_upd
            sql_upd="select id from Class where number = "&request("usclassno")&" and specialtyID = "&request("usspe")
            set rec_upd = conn.execute(sql_upd)
            if rec_upd.eof then
                response.write "<script>alert(""没有找到"&request("class")&"班级"")</script>"
            else
                classid=rec_upd(0)
                sql_upd_exec=" update Student set name='"&request("usname")&"', gender="&request("usgender")&", specialtyID="&request("usspe")&", classID="&classid&", grade="&request("usgrade")&", birth='"&request("usbirth")&"' where id =" &request("usno")
                'response.write sql_upd_exec
                'response.write "<p>消息：添加成功!</p>"
                conn.execute sql_upd_exec
                '语句原型 "update Student set name='a', gender=1, specialtyID=1, classID=1, grade=1, birth='' where id = 1"
            end if
        end if
    %>
    <form>
        <div>
            <label>学号：</label><input type="text" size="7" name="sno" id="sno" value="<%=request("sno") %>" />
        </div>
        <div>
            <label>姓名：</label><input type="text" size="7" name="sname" id="sname" value="<%=request("sname") %>" />
        </div>
        <div>
            <label>性别：</label>
            <select name="sgender" id="sgender">
                <option id="sg" value="" <%if request("sgender")="" then %>selected="selected" <%end if %>></option>
                <option id="sg0" value="0" <%if request("sgender")="0" then %>selected="selected" <%end if %>>女</option>
                <option id="sg1" value="1" <%if request("sgender")="1" then %>selected="selected" <%end if %>>男</option>
            </select>
        </div>
        <div>
            <label>专业：</label>
            <select name="sspe" id="sspe">
                <option id="ss" value="" <%if request("sspe")=""      then %>selected="selected" <%end if %>></option>
                <option id="ss1" value="10001" <%if request("sspe")="10001" then %>selected="selected" <%end if %>>软件工程</option>
                <option id="ss2" value="10002" <%if request("sspe")="10002" then %>selected="selected" <%end if %>>数字媒体与技术</option>
                <option id="ss3" value="10003" <%if request("sspe")="10003" then %>selected="selected" <%end if %>>计算机科学与技术</option>
                <option id="ss4" value="10004" <%if request("sspe")="10004" then %>selected="selected" <%end if %>>机械设计制造及其自动化</option>
                <option id="ss5" value="10005" <%if request("sspe")="10005" then %>selected="selected" <%end if %>>自动化</option>
                <option id="ss6" value="10006" <%if request("sspe")="10006" then %>selected="selected" <%end if %>>测控技术与仪器</option>
                <option id="ss7" value="10007" <%if request("sspe")="10007" then %>selected="selected" <%end if %>>电子信息科学与技术</option>
                <option id="ss8" value="10008" <%if request("sspe")="10008" then %>selected="selected" <%end if %>>通信工程</option>
            </select>
        </div>
        <div>
            <label>班级：</label>
            <select name="sclass" id="sclass">
                <option id="sc" value="" <%if request("sclass")=""  then %>selected="selected" <%end if %>></option>
                <option id="sc1" value="1" <%if request("sclass")="1" then %>selected="selected" <%end if %>>1</option>
                <option id="sc2" value="2" <%if request("sclass")="2" then %>selected="selected" <%end if %>>2</option>
                <option id="sc3" value="3" <%if request("sclass")="3" then %>selected="selected" <%end if %>>3</option>
            </select>
        </div>
        <div>
            <label>年级：</label><input type="number" value="<%=request("sgrade")%>" name="sgrade" id="sgrade" />
        </div>
        <div>
            <input type="submit" value="筛选" />
            <input type="reset" value="清空" onclick="myResetBtn();"/>
        </div>
    </form>
    <hr width="100%" />
    <div class="table">
        <table border="1" width="107%">
            <%
                '准备数据
                rsno=request("sno")        &""
                rsname=request("sname")    &""
                rsgender=request("sgender")&""
                rsspe=request("sspe")      &""
                rsclass=request("sclass")  &""
                rsgrade=request("sgrade")  &""
                sqlStr="select Student.ID 学号,Student.name 姓名,Student.gender 性别,Specialty.name 专业, Class.name 班级,Student.grade 年级,Student.birth 出生年月, Class.ID, Specialty.ID from Student,Specialty,Class where Student.specialtyID=Specialty.ID and Student.classID=Class.ID"

                '附加条件
                if rsno<>"" then
                    sqlStr=sqlStr&" and Student.ID = "&rsno
                end if
                if rsname<>"" then
                    sqlStr=sqlStr&" and Student.name like '%"&rsname&"%'"
                end if
                if rsgender<>"" then
                    sqlStr=sqlStr&" and Student.gender = "&rsgender
                end if
                if rsspe<>"" then
                    sqlStr=sqlStr&" and Specialty.id = "&rsspe
                end if
                if rsclass<>"" then
                    sqlStr=sqlStr&" and Class.number = "&rsclass
                end if
                if rsgrade<>"" then
                    sqlStr=sqlStr&" and Student.grade = "&rsgrade
                end if
                
                'response.write sqlStr

                '执行语句
                if request("sql")&""<>"" then 
                    set rec=conn.execute(request("sql"))
                else
                    set rec=Conn.execute(sqlStr)
                end if
            %>
            <tr height="40px">
                <%
                '循环输出题头 
                for i=0 to 6 'rec.fields.count-3
                    response.write "<th>"&rec.fields(i).name&"</th>"
                next
                response.write "<th colspan='2' align='center'><a href='insertStu.asp'>添加学生</a></th>"
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
                do while not rec.eof
                    '每行记录的开始
                    response.write "<tr height=""40px"">"
                
                    '若当前行不是需要修改的行 
                    if id_forUpd&""<>ii&"" then
                        for i=0 to 6 'rec.fields.count-3
                            if i=2 then
                                if rec(i)=0 then
                                    response.Write("<td>女</td>")
                                else
                                    if rec(i)=1 then
                                        response.Write("<td>男</td>")
                                    end if
                                end if
                            else
	                            response.write "<td>"&rec(i)&"</td>"
                            end if
	                    next
                        if request("upd")&""<>"1" then
                            response.write "<td><a href='StudentInfo.asp?upd=1&count="&ii&"&sql="&sqlStr&"'>修改</a></td>"
                            'response.write "<td><input type='button' onclick=""window.location.href='StudentInfo.asp?upd=1&count="&ii&"&sql="&sqlStr&"';return false;"" value='修改' /></td>"
                            response.write "<td><a href='StudentInfo.asp?del=1&id="&rec(0)&"&sname="&request("sname")&"&sgender="&request("sgender")&"&sspe="&request("sspe")&"&sclass="&request("sclass")&"&sgrade="&request("sgrade")&"'>删除</a></td>"
                        end if
                    else'若是需要修改的行
                        '准备数据
                        dim rec_class
                        set rec_class = conn.execute("select Class.number from Class where ID = "&rec(7))
                        if rec_class.eof then
                            response.write "<script>alert'没有这个班级'</script>"
                        else
                            class_number=rec_class(0)&""
                        end if
                        '输出表格
                        response.write "<form name='update' >"
                        response.write "<input type='hidden' name='usno' value='"&rec(0)&"' />"
                        response.write "<td>"&rec(0)&"</td>"
                        response.write "<td><input type='text' size='7' name='usname' value='"&rec(1)&"' /></td>"
                        response.write "<td><select name='usgender'>"
                            response.write "<option value='0' " : if rec(2)="0" then : response.write "selected='selected'" : end if : response.write ">女</option>"
                            response.write "<option value='1' " : if rec(2)="1" then : response.write "selected='selected'" : end if : response.write ">男</option>"
                        response.write "</select></td>"
                        response.write "<td><select name='usspe'>"
                            response.write "<option value='10001' " : if rec(8)="10001" then : response.write "selected='selected'" : end if : response.write">软件工程</option>"
                            response.write "<option value='10002' " : if rec(8)="10002" then : response.write "selected='selected'" : end if : response.write">数字媒体与技术</option>"
                            response.write "<option value='10003' " : if rec(8)="10003" then : response.write "selected='selected'" : end if : response.write">计算机科学与技术</option>"
                            response.write "<option value='10004' " : if rec(8)="10004" then : response.write "selected='selected'" : end if : response.write">机械设计制造及其自动化</option>"
                            response.write "<option value='10005' " : if rec(8)="10005" then : response.write "selected='selected'" : end if : response.write">自动化</option>"
                            response.write "<option value='10006' " : if rec(8)="10006" then : response.write "selected='selected'" : end if : response.write">测控技术与仪器</option>"
                            response.write "<option value='10007' " : if rec(8)="10007" then : response.write "selected='selected'" : end if : response.write">电子信息科学与技术</option>"
                            response.write "<option value='10008' " : if rec(8)="10008" then : response.write "selected='selected'" : end if : response.write">通信工程</option>"
                        response.write "</select></td>"
                        response.write "<input type='hidden' name='usclassid' value='"&rec(7)&"' />"
                        response.write "<td><select name='usclassno'>"
                            response.write "<option value='1' " : if class_number="1" then : response.write "selected='selected'" : end if : response.write ">1</option>"
                            response.write "<option value='2' " : if class_number="2" then : response.write "selected='selected'" : end if : response.write ">2</option>"
                            response.write "<option value='3' " : if class_number="3" then : response.write "selected='selected'" : end if : response.write ">3</option>"
                        response.write "</select></td>"
                        response.write "<td><input type='number' value='"&rec(5)&"' name='usgrade' /></td>"
                        response.write "<td><input type='date' value='"&rec(6)&"' name='usbirth' /></td>"

                        'response.write "<td><input type='button' value='提交' onclick=""window.location.href='StudentInfo.asp?exec_upd=1&id="&rec(0)&"&name="&rec(1)&"&stugender="&rec(2)&"&spe="&rec(3)&"&class="&rec(4)&"&grade="&rec(5)&"&birth="&rec(6)&"';return false;""/></td>"
                        response.write "<td><input type='button' value='取消' onclick=""window.location.href='StudentInfo.asp?sql="&request("sql")&"';return false;""/></td>"
                        response.write "<td><input type='submit' name='exec_upd_Btn' value='提交' /></td>"
                    
                        response.write "</form>"
                    end if

	                response.write "</tr>"
	                rec.movenext
                    ii=ii+1
	            loop
            %>
        </table>
    </div>

    <%
        '释放资源
        rec.close
        set rec=nothing
        Conn.close
        set Conn=nothing
    %>
</body>
