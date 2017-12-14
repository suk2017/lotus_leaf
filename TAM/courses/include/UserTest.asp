<%
    if session("userid")&""="" then
        response.redirect "http://127.0.0.1/out/blank.html"
	    'response.Write("<script>window.open('>http://127.0.0.1/out/login.asp?info=ÇëµÇÂ¼')</script>")
    end if
    if session("user")&""<>"tamanager" then
        response.redirect "http://127.0.0.1/out/blank.html"
	    'response.Write("<script>window.open('>http://127.0.0.1/out/login.asp?info=ÇëµÇÂ¼')</script>")
    end if
%>