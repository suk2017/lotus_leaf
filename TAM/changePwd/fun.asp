<!--#include file="include/connect.asp"-->
<!--#include file="include/UserTest.asp"-->
<head>
    <title>�޸�����</title>
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
    '�ж������ʽ�Ƿ���ȷ
    if request("oldPassword")&"" = "" then
        response.redirect "changePassword.asp?info=ԭ�����������"
    elseif request("newPassword")&"" = "" then
        response.Redirect("changePassword.asp?info=�����������룡")
    elseif request("newPassword2")&"" = "" then
        response.Redirect("changePassword.asp?info=���ٴ����������룡")
    elseif request("newPassword") <> request("newPassword2") then
        response.Redirect("changePassword.asp?info=������������������һ�£�")
    end if 
    
    '�ж������Ƿ���ȷ
    set rec = conn.execute("select ID from TAManager where ID = "& session("id") & " and password = '" & request("oldPassword") & "'")
    if rec.eof then
        response.redirect "changePassword.asp?info=������֤����"
    else
        '������֤��ȷ
        set rec2 = conn.execute("update TAManager set password = '" & request("newPassword") & "' where ID = " & session("id"))
        response.Redirect "changePassword.asp?info=�޸ĳɹ���"
    end if

    rec.close
    Set rec = Nothing 
    rec2.close
    set rec2 = Nothing
    Conn.close
    Set Conn = Nothing
%>