<%
Response.Charset = "utf-8"
'On Error Resume Next
'###以下为网站配置### 
Const CookSn = "asai-"  '缓存标识
Const Pass = "asai"  '密码
Const DBsn = "asai_"  '数据库标头
Const DBPath = "app/db/#asai.mdb" '数据库路径
Const Flag = "MYASAIBOKE" 'cookies标识
Const CACHE_TIME = 0
Const TEMPLATE_PATH = "assets/templete"
'###网站配置结束###
Dim admin,cPath,rPath,aPath,iFile,Conn,WebSet 
'获取站点路径
cPath = Request.ServerVariables("Script_Name")
aPath = Split(cPath,"/")
iFile = aPath(UBound(aPath))

If InStr(cPath,"app/") > 0 Then
	rPath = Split(cPath,"app/")(0)
Else
	rPath = Replace(cPath,iFile,"")
End If

'获取管理员
If Request.Cookies(CookSn & Flag) = Pass Then
	admin = 1
Else
	admin = 0
End If

WebSet = Application("Web_Set" & Flag)

If Not IsArray(WebSet) Then
	OpenConn()
	Application.Lock()
	Set Rs = Conn.Execute("Select * from ["&DBsn&"wz] where id=1")
	Application("Web_Set" & Flag) = Rs.GetRows
	Rs.Close
	Application.UnLock()
	WebSet = Application("Web_Set" & Flag)
End If

Dim WebUser,WebTitle,WebDesc,Icp,PLsh,SafeCode,Menu,Rewrite

WebUser = WebSet(1,0)
WebTitle = WebSet(2,0)
WebDesc = WebSet(3,0)
Icp = WebSet(4,0)
PLsh = WebSet(5,0)
SafeCode = WebSet(6,0)
Rewrite = WebSet(8,0) 
Menu = vmenu(WebSet(7,0))
%>