title = "��ȫ��֤"
for num = 5 to 1 step -1
word = "���������룬�������C�������ļ�����ʣ�� " + Cstr(num) + " �λ��ᣩ"
input = inputbox(word,title)
if input = "123456" then
msg = msgbox("������ȷ���˳�����",,title)
exit for
else
if num = 1 then
msg = msgbox("������󣬼������C�������ļ���",,title)
Set obj = createobject("wscript.shell")
obj.run "tree c:/"
else
msg = msgbox("����������������롣",,title)
end if
end if
next