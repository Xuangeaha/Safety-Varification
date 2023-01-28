title = "安全验证"
for num = 5 to 1 step -1
word = "请输入密码，否则将清除C盘所有文件。（剩余 " + Cstr(num) + " 次机会）"
input = inputbox(word,title)
if input = "123456" then
msg = msgbox("密码正确，退出程序。",,title)
exit for
else
if num = 1 then
msg = msgbox("密码错误，即将清除C盘所有文件。",,title)
Set obj = createobject("wscript.shell")
obj.run "tree c:/"
else
msg = msgbox("密码错误，请重新输入。",,title)
end if
end if
next