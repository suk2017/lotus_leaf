<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>修改密码</title>
    <meta charset="gb2312" />
</head>
<%
    'if session("type")&"" <> "TAManager" then
    '    response.Redirect ""
    'end if
    'if session("id")&"" = "" then
    '    response.Redirect ""
    'end if
    session("id")=10001
%>
<%
    '判断输入格式是否正确
    if request("oldPassword")&"" = "" then
        response.redirect "changePassword.asp?info=原密码输入错误！"
    elseif request("newPassword")&"" = "" then
        response.Redirect("changePassword.asp?info=请输入新密码！")
    elseif request("newPassword2")&"" = "" then
        response.Redirect("changePassword.asp?info=请再次输入新密码！")
    elseif request("newPassword") <> request("newPassword2") then
        response.Redirect("changePassword.asp?info=两次输入的新密码必须一致！")
    end if 
    
    '判断密码是否正确
    set rec = conn.execute("select ID from TAManager where ID = "& session("id") & " and password = '" & request("oldPassword") & "'")
    if rec.eof then
        response.redirect "changePassword.asp?info=密码验证错误！"
    else
        '密码验证正确
        set rec2 = conn.execute("update TAManager set password = '" & request("newPassword") & "' where ID = " & session("id"))
        response.Redirect "changePassword.asp?info=修改成功！"
    end if

    rec.close
    Set rec = Nothing 
    rec2.close
    set rec2 = Nothing
    Conn.close
    Set Conn = Nothing
%>