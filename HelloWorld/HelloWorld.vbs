dim return
return = msgbox("hello world!!", vbokcancel)

if return = vbcancel then
   msgbox "キャンセルボタンが押下されました。", vbokonly 
else
   msgbox "ＯＫボタンが押下されました。", vbokonly 
end if



