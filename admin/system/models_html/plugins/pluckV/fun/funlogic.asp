<%  
'======================================================================================
'C9 ��̬���·���ϵͳ
'�ٷ���վ:http://www.csc9.cn
'======================================================================================
'�ж������Ƿ���ڴ��ַ�
'str1:���� str2:�ַ�
function boolstr(str1,str2)
  dim arrstr
  arrstr=split(str1,str2)
  if ubound(arrstr)>0 then
    boolstr="1"
  else
    boolstr="0"
  end if
end function
'==============================
'ȡ���м����ݺ��� midstr(str,str1,str2)
'str ���� str1,��ʼ str2,����
'==============================
function midstr(str,str1,str2)
  dim midstr1,midstr2
  midstr1=split(str,str1)
  if ubound(midstr1)>0 then
    midstr2=split(midstr1(1),str2)
	if ubound(midstr2)>0 then
	  midstr=midstr2(0)
	else
	  midstr="0"
	end if
  else
    midstr="0"
  end if
end function

'==============================
'��ȡ�ַ�����GetStrLen
'����strΪָ���ַ�,����Ϊ2. a,Ϊ�������
'==============================
Function GetStrLen(str,a)
If IsNull(str) Or str = "" Then
getStrLen = 0
Else
Dim i, n, k, chrA, chrb
k = 0
n = Len(str)
For i = 1 To n
chrA = Mid(str, i, 1)
If Asc(chrA) >= 0 And Asc(chrA) <= 255 Then
k = k + 1
Else
k = k + 2
End If

chrb=chrb & chrA

if k>=a then
exit for
end if

Next
'�ڴ����K�ɵ��ַ�����,����Ϊ2
getStrLen = chrb
End If
End Function

'==================================================
'�ַ����� funstr
'2009-06-05 Crazy
'==================================================
Function funstr(str)	 
	str = trim(str) 	 
	str = replace(str, "<", "&lt;", 1, -1, 1)
	str = replace(str, ">", "&gt;", 1, -1, 1)
	str = replace(str,"'","��")
	
	funstr = str
End Function

Function unfunstr(str)	 	 
	str = replace(str, "&lt;","<", 1, -1, 1)
	str = replace(str, "&gt;",">", 1, -1, 1)
	str = trim(str)
	str = replace(str,"��","'")	
	unfunstr = str
End Function
'==================================================
'��ȡ�ļ�����
'==================================================
function openfile(url)
  fileurl=server.MapPath(url)
  set fso=server.CreateObject("scripting.filesystemobject") '����FSO
  set mofile=fso.opentextfile(fileurl,1) '�Զ��ķ�ʽ���ļ�
  mo_top=mofile.readall() '��ȡȫ������
  mofile.close
  
  openfile=mo_top
end function
'==================================================
'д���ļ�����
'==================================================
sub createfile(url,str)
  fileurl=server.MapPath(url)
  set fso=server.CreateObject("scripting.filesystemobject") 
  set mofile=fso.createtextfile(fileurl,true)
  mofile.write str
  mofile.close

end sub
'==================================================
'����ļ����Ƿ����,����������򴴽�
'==================================================
sub createfolder(folder)
  folderurl=server.MapPath(folder)
  set fso=server.CreateObject("scripting.filesystemobject") 
  if fso.folderexists(folderurl) then
  
  else
    fso.createfolder(folderurl)
  end if
set fso=nothing '''*****
end sub

'==================================================
'ɾ���ļ��� 
'==================================================
sub delfolder(folder)
  folderurl=server.MapPath(folder)
  set fso=server.CreateObject("scripting.filesystemobject") 
  if fso.folderexists(folderurl) then
    fso.deletefolder folderurl,true
  end if
set fso=nothing '''*****
end sub
'==================================================
'ɾ���ļ�
'==================================================
sub delfile(files)
  fileurl=server.MapPath(files)
  set fso=server.CreateObject("scripting.filesystemobject") '����FSO
  if fso.fileexists(fileurl) then
    fso.deletefile fileurl,true
  end if
set fso=nothing '''*****
end sub

'==================================================
'����ļ� ����ֵ ��:1 ��:0
'==================================================
function fileex(files)
  fileurl=server.MapPath(files)
  set fso=server.CreateObject("scripting.filesystemobject") '����FSO
  if fso.fileexists(fileurl) then
    fileex=1
  else
    fileex=0
  end if
set fso=nothing '''*****
end function
'==================================================
'�����Ի���
'==================================================
sub getshow(str,url)
  if str="" and url <>"" then
    response.Write "<script language='javascript'>window.document.location.href='"& url &"'</script>"
    response.End()
  end if
  if str<>"" and url="" then
    response.Write "<script language='javascript'>alert('"&str&"');history.go(-1)</script>"
    response.End()
  end if
  if str<>"" and url<>"" then
    response.Write "<script language='javascript'>alert('"&str&"');window.document.location.href='"& url &"'</script>"
    response.End()
  end if
  
end sub
'==================================================
'����������
'==================================================
function showtemp()
  set fso=server.CreateObject("scripting.filesystemobject")
  set ff=fso.getfolder(server.MapPath("../temp"))
  set openfs=ff.subfolders
  for each openf in openfs
  showtemp=showtemp&"<option value='"&openf.name&"'>"&openf.name&"</option>"
  next
  set fso=nothing
  set ff=nothing
  set openfs=nothing
end function
'==================================================
'�����ϴ��ļ������� filenamestr
'==================================================
function filenamestr()
   str=now()
   strs=split(str," ")
   strsa=split(strs(0),"-")
   strsb=split(strs(1),":")
   filenamestr=""
   filenamestr=filenamestr&strsa(0)
   filenamestr=filenamestr&strsa(1)
   filenamestr=filenamestr&strsa(2)
   filenamestr=filenamestr&strsb(0)
   filenamestr=filenamestr&strsb(1)
   filenamestr=filenamestr&strsb(2)
end function

function filenametime(str)
  arr=split(str,"-")
  filenametime=arr(0)&arr(1)&arr(2)
end function

function timemd(str)
  arr=split(str,"-")
  if len(arr(2))=1 then
    arr(2)="0"&arr(2)
  end if
  if len(arr(1))=1 then
    arr(1)="0"&arr(1)
  end if
  timemd=arr(1)&"-"&arr(2)
end function
'==================================================
'���HTML���뺯��
'==================================================
Function ClearHtml(Content) 
Content=Zxj_ReplaceHtml("&#[^>]*;", "", Content) 
Content=Zxj_ReplaceHtml("</?marquee[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?object[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?param[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?embed[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?table[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml(" ","",Content) 
Content=Zxj_ReplaceHtml("</?tr[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?th[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?p[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?a[^>]*>","",Content) 
'Content=Zxj_ReplaceHtml("</?img[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?tbody[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?li[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?span[^>]*>","",Content) 

Content=Zxj_ReplaceHtml("</?div[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?th[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?td[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?script[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("(javascript|jscript|vbscript|vbs):", "", Content) 
Content=Zxj_ReplaceHtml("on(mouse|exit|error|click|key)", "", Content) 
Content=Zxj_ReplaceHtml("<\\?xml[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("<\/?[a-z]+:[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?font[^>]*>", "", Content) 
Content=Zxj_ReplaceHtml("</?b[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?u[^>]*>","",Content) 
Content=Zxj_ReplaceHtml("</?i[^>]*>","",Content)
Content=Zxj_ReplaceHtml("</?strong[^>]*>","",Content) 

ClearHtml=Content 
End Function 

Function Zxj_ReplaceHtml(patrn, strng,content) 
IF IsNull(content) Then 
content="" 
End IF 
Set regEx = New RegExp ' ����������ʽ�� 
regEx.Pattern = patrn ' ����ģʽ�� 
regEx.IgnoreCase = true ' ���ú����ַ���Сд�� 
regEx.Global = true ' ����ȫ�ֿ����ԡ� 
Zxj_ReplaceHtml=regEx.Replace(content,strng) ' ִ������ƥ�� 
End Function

'ƥ���һ����ϵ�Ҫ���ֵ���

function reimgone(patrn,str)
  reimgone=""
  Set re = New RegExp
  re.Pattern = patrn
  re.IgnoreCase = true
  re.Global = true
  set reexe=re.Execute(str)
  
    if isnull(reexe(0)) then
      reimgone=""
    else
	  reimgone=reexe(0)
    end if
  
end function

'��ȡ�������ݺ���
 Function Gethttppage(Path,crzchars)      
 T = Getbody(Path)
 Gethttppage=Bytestobstr(T,crzchars)
 End Function

 Function Getbody(Url)      
 On Error Resume Next
 Set Retrieval = Createobject("Microsoft.Xmlhttp") 
     Retrieval.Open "Get",Url, False, "", "" 
     Retrieval.send()
     Getbody = Retrieval.Responsebody
 Set Retrieval = Nothing 
 End Function 

Function BytesToBstr(body,Cset)         
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
        objstream.Type = 1
        objstream.Mode =3
        objstream.Open
        objstream.Write body
        objstream.Position = 0
        objstream.Type = 2
        objstream.Charset = Cset
        BytesToBstr = objstream.ReadText 
        objstream.Close
        set objstream = nothing
End Function

Function Newstring(wstr,strng)              '
        Newstring=Instr(lcase(wstr),lcase(strng))
        if Newstring<=0 then Newstring=Len(wstr)
End Function
'��ȡ�������ݺ�������


'Crazy�������º��� 2009-06-06
'==================================================
'��������CheckDir2
'�� �ã�����ļ����Ƿ����
'�� ����FolderPath ------�ļ��е�ַ
'==================================================
Function CheckDir2(byval FolderPath)
dim fso
folderpath=Server.MapPath(".")&"\"&folderpath
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If fso.FolderExists(FolderPath) then
'����
  CheckDir2 = True
Else
'������
  CheckDir2 = False
End if
Set fso = nothing
End Function


'==================================================
'��������MakeNewsDir2
'�� �ã������µ��ļ���
'�� ����foldername ------�ļ�������
'==================================================
Function MakeNewsDir2(byval foldername)
dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
fso.CreateFolder(Server.MapPath(".") &"\" &foldername)
If fso.FolderExists(Server.MapPath(".") &"\" &foldername) Then
MakeNewsDir2 = True
Else
MakeNewsDir2 = False
End If
Set fso = nothing
End Function
'==================================================
'��������DefiniteUrl
'�� �ã�����Ե�ַת��Ϊ���Ե�ַ
'�� ����PrimitiveUrl ------Ҫת������Ե�ַ
'�� ����ConsultUrl ------��ǰ��ҳ��ַ
'==================================================
Function DefiniteUrl(Byval PrimitiveUrl,Byval ConsultUrl)
Dim ConTemp,PriTemp,Pi,Ci,PriArray,ConArray
If PrimitiveUrl="" or ConsultUrl="" or PrimitiveUrl="$False$" Then
DefiniteUrl="$False$"
Exit Function
End If
If Left(ConsultUrl,7)<>"HTTP://" And Left(ConsultUrl,7)<>"http://" Then
ConsultUrl= "http://" & ConsultUrl
End If
ConsultUrl=Replace(ConsultUrl,"://",":\\")
If Right(ConsultUrl,1)<>"/" Then
If Instr(ConsultUrl,"/")>0 Then
If Instr(Right(ConsultUrl,Len(ConsultUrl)-InstrRev(ConsultUrl,"/")),".")>0 then 
Else
ConsultUrl=ConsultUrl & "/"
End If
Else
ConsultUrl=ConsultUrl & "/"
End If
End If
ConArray=Split(ConsultUrl,"/")
If Left(PrimitiveUrl,7) = "http://" then
DefiniteUrl=Replace(PrimitiveUrl,"://",":\\")
ElseIf Left(PrimitiveUrl,1) = "/" Then
DefiniteUrl=ConArray(0) & PrimitiveUrl
ElseIf Left(PrimitiveUrl,2)="./" Then
DefiniteUrl=ConArray(0) & Right(PrimitiveUrl,Len(PrimitiveUrl)-1)
ElseIf Left(PrimitiveUrl,3)="../" then
Do While Left(PrimitiveUrl,3)="../"
PrimitiveUrl=Right(PrimitiveUrl,Len(PrimitiveUrl)-3)
Pi=Pi+1
Loop 
For Ci=0 to (Ubound(ConArray)-1-Pi)
If DefiniteUrl<>"" Then
DefiniteUrl=DefiniteUrl & "/" & ConArray(Ci)
Else
DefiniteUrl=ConArray(Ci)
End If
Next
DefiniteUrl=DefiniteUrl & "/" & PrimitiveUrl
Else
If Instr(PrimitiveUrl,"/")>0 Then
PriArray=Split(PrimitiveUrl,"/")
If Instr(PriArray(0),".")>0 Then
If Right(PrimitiveUrl,1)="/" Then
DefiniteUrl="http:\\" & PrimitiveUrl
Else
If Instr(PriArray(Ubound(PriArray)-1),".")>0 Then 
DefiniteUrl="http:\\" & PrimitiveUrl
Else
DefiniteUrl="http:\\" & PrimitiveUrl & "/"
End If
End If 
Else
If Right(ConsultUrl,1)="/" Then 
DefiniteUrl=ConsultUrl & PrimitiveUrl
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & PrimitiveUrl
End If
End If
Else
If Instr(PrimitiveUrl,".")>0 Then
If Right(ConsultUrl,1)="/" Then
If right(PrimitiveUrl,3)=".cn" or right(PrimitiveUrl,3)="com" or right(PrimitiveUrl,3)="net" or right(PrimitiveUrl,3)="org" Then
DefiniteUrl="http:\\" & PrimitiveUrl & "/"
Else
DefiniteUrl=ConsultUrl & PrimitiveUrl
End If
Else
If right(PrimitiveUrl,3)=".cn" or right(PrimitiveUrl,3)="com" or right(PrimitiveUrl,3)="net" or right(PrimitiveUrl,3)="org" Then
DefiniteUrl="http:\\" & PrimitiveUrl & "/"
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl
End If
End If
Else
If Right(ConsultUrl,1)="/" Then
DefiniteUrl=ConsultUrl & PrimitiveUrl & "/"
Else
DefiniteUrl=Left(ConsultUrl,InstrRev(ConsultUrl,"/")) & "/" & PrimitiveUrl & "/"
End If 
End If
End If
End If
If Left(DefiniteUrl,1)="/" then
DefiniteUrl=Right(DefiniteUrl,Len(DefiniteUrl)-1)
End if
If DefiniteUrl<>"" Then
DefiniteUrl=Replace(DefiniteUrl,"//","/")
DefiniteUrl=Replace(DefiniteUrl,":\\","://")
Else
DefiniteUrl="$False$"
End If
End Function
'==================================================
'��������ReplaceSaveRemoteFile
'�� �ã��滻������Զ���ļ�
'�� ����ConStr ------ Ҫ�滻���ַ���
'�� ����StarStr ----- ǰ��
'�� ����OverStr ----- 
'�� ����IncluL ------ 
'�� ����IncluR ------ 
'�� ����SaveTf ------ �Ƿ񱣴��ļ���False�����棬True����
'�� ����SaveFilePath- �����ļ���
'�� ��: TistUrl------ ��ǰ��ҳ��ַ
'==================================================
Function ReplaceSaveRemoteFile(ConStr,StartStr,OverStr,IncluL,IncluR,SaveTf,SaveFilePath,TistUrl)
If ConStr="$False$" or ConStr="" Then
ReplaceSaveRemoteFile="$False$"
Exit Function
End If
Dim TempStr,TempStr2,ReF,Matches,Match,Tempi,TempArray,TempArray2,OverTypeArray

Set ReF = New Regexp 
ReF.IgnoreCase = True 
ReF.Global = True
ReF.Pattern = "("&StartStr&").+?("&OverStr&")"
Set Matches =ReF.Execute(ConStr) 
For Each Match in Matches
If Instr(TempStr,Match.Value)=0 Then
If TempStr<>"" then 
TempStr=TempStr & "$Array$" & Match.Value
Else
TempStr=Match.Value
End if
End If
Next 
Set Matches=nothing
Set ReF=nothing
If TempStr="" or IsNull(TempStr)=True Then
ReplaceSaveRemoteFile=ConStr
Exit function
End if
If IncluL=False then
TempStr=Replace(TempStr,StartStr,"")
End if
If IncluR=False then
If Instr(OverStr,"|")>0 Then
OverTypeArray=Split(OverStr,"|")
For Tempi=0 To Ubound(OverTypeArray) 
TempStr=Replace(TempStr,OverTypeArray(Tempi),"")
Next
Else
TempStr=Replace(TempStr,OverStr,"")
End If 
End if
TempStr=Replace(TempStr,"""","")
TempStr=Replace(TempStr,"'","")

Dim RemoteFile,RemoteFileurl,SaveFileName,SaveFileType,ArrSaveFileName,RanNum
If Right(SaveFilePath,1)="/" then
SaveFilePath=Left(SaveFilePath,Len(SaveFilePath)-1)
End If
If SaveTf=True then
If CheckDir2(SaveFilePath)=False Then
If MakeNewsDir2(SaveFilePath)=False Then
SaveTf=False
End If
End If
End If
SaveFilePath=SaveFilePath & "/"

'ͼƬת��/����
TempArray=Split(TempStr,"$Array$")
For Tempi=0 To Ubound(TempArray)
RemoteFileurl=DefiniteUrl(TempArray(Tempi),TistUrl)
If RemoteFileurl<>"$False$" And SaveTf=True Then'����ͼƬ
  ArrSaveFileName = Split(RemoteFileurl,".")
  SaveFileType=ArrSaveFileName(Ubound(ArrSaveFileName))'�ļ�����
  RanNum=Int(900*Rnd)+100
  SaveFileName = SaveFilePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&SaveFileType  
  Call SaveRemoteFile(SaveFileName,RemoteFileurl)
ConStr=Replace(ConStr,TempArray(Tempi),SaveFileName)
ElseIf RemoteFileurl<>"$False$" and SaveTf=False Then'������ͼƬ
SaveFileName=RemoteFileUrl
ConStr=Replace(ConStr,TempArray(Tempi),SaveFileName)
End If
If RemoteFileUrl<>"$False$" Then
If UploadFiles="" then
UploadFiles=SaveFileName
Else
UploadFiles=UploadFiles & "|" & SaveFileName
End if
End If
Next 
ReplaceSaveRemoteFile=ConStr
End function
'==================================================
'��������SaveRemoteFile
'�� �ã�����Զ�̵��ļ�������
'�� ����LocalFileName ------ �����ļ���
'�� ����RemoteFileUrl ------ Զ���ļ�URL
'==================================================
Function SaveRemoteFile(LocalFileName,RemoteFileUrl)
dim Ads,Retrieval,GetRemoteData
Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
With Retrieval
  .Open "Get", RemoteFileUrl, False, "", ""
  .Send
  GetRemoteData = .ResponseBody
End With
Set Retrieval = Nothing
Set Ads = Server.CreateObject("Adodb.Stream")
With Ads
  .Type = 1
  .Open
  .Write GetRemoteData
  .SaveToFile server.MapPath(LocalFileName),2
  .Cancel()
  .Close()
End With
Set Ads=nothing
'��ʼ��ˮӡ
If IsObjInstalled("Persits.Jpeg") then
'call Shuiyin(LocalFileName,18,"Arial",18,18,Crzshuiyin,1,Crzshuiyincolor) '����Crzshuiyin
call Shuiyin(LocalFileName,Crzshuiyinsize,Crzshuiyinfamily,Crzshuiyinleft,Crzshuiyinright,Crzshuiyin,0,Crzshuiyincolor) '����Crzshuiyin
End If
end Function

'==================================================
'��������GetImg
'�� �ã�ȡ�������е�һ��ͼƬ
'�� ����str ------ ��������
'�� ����strpath ------ ����ͼƬ��·��
'==================================================
Function GetImg(str,strpath)
set objregEx = new RegExp
objregEx.IgnoreCase = true
objregEx.Global = true
zzstr=""&strpath&"(.+?)\.(jpg|gif|png|bmp)"
objregEx.Pattern = zzstr
set matches = objregEx.execute(str)
for each match in matches
retstr = retstr &"|"& Match.Value
next
if retstr<>"" then
Imglist=split(retstr,"|")
Imgone=replace(Imglist(1),strpath,"")
GetImg=Imgone
else
GetImg=""
end if
end function


'���ˮӡ����
Function Shuiyin(imgurl,fontsize,family,top,left,content,Horflip,Color) '���ù����� 
'����call Shuiyin("c9cms.jpg",13,"����",18,18,"www.c9cms.cn",1) 
Dim Jpeg,font_size,font_family,f_content,f_Horflip 
'����ʵ�� 
Set Jpeg = Server.CreateObject("Persits.Jpeg") 
font_size=10 
font_family="����" 
f_left= 5 
f_top=5 
if imgurl<>"" then 
Jpeg.Open Server.MapPath(imgurl)'ͼƬ·�������� 
else 
response.write "δ�ҵ�ͼƬ·��" 
exit Function 
end if 
if fontsize<>"" then font_size=fontsize '�����С 
if family<>"" then font_family=family '���� 
if top<>"" then f_left=left 'ˮӡ��ͼƬ���λ�� 
if left<>"" then f_top=top 'ˮӡ��ͼƬtopλ�� 
if content="" then 'ˮӡ���� 
response.write "ˮӡʲô�����أ�ˮӡ���ɹ���" 
exit Function 
else 
f_content=content 
end if 
' �������ˮӡ 
Jpeg.Canvas.Font.Color = Color  'ff0000     '��ɫ
Jpeg.Canvas.Font.Family = font_family '����
jpeg.canvas.font.size= font_size    '�������
Jpeg.Canvas.Font.Bold = True        '�Ӵ�
Jpeg.Canvas.Font.Quality = 4        '�������
If Horflip = 1 Then 
Jpeg.FlipH 'ͼƬ������
'Jpeg.SendBinary 
End If 
Jpeg.Canvas.Print f_left, f_top, f_content 
' �����ļ� 
Jpeg.Save Server.MapPath(imgurl) 
' ע������ 
Set Jpeg = Nothing 
'response.write "ˮӡ�ɹ���ͼƬ�ϼ��� "&content&"" 
End Function 


'**************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'        False ----û�а�װ
'**************************************************
Function IsObjInstalled(strClassString)
 IsObjInstalled = False
 Err = 0
 Dim xTestObj
 Set xTestObj = Server.CreateObject(strClassString)
 If 0 = Err Then IsObjInstalled = True
 Set xTestObj = Nothing
 Err = 0
End Function



'*********************************************************
'������ DetectUrl
'���ܣ� �滻�ַ����е�Զ���ļ����·��Ϊ��http://..��ͷ�ľ���·��
'������ sContent Ҫ����ĺ����·������ҳ���ı�����
'sUrl�� �������Զ����ҳ�����URL�����ڷ������·��
'���أ� �滻�������Ϊ��������֮����µ���ҳ�ı�����
'*********************************************************
Function DetectUrl(sContent,sUrl)
Dim re,sMatch
Set re=new RegExp
re.Multiline=True
re.IgnoreCase =true
re.Global=True
re.Pattern = "(http://[-A-Z0-9.]+)/[-A-Z0-9+&@#%~_|!:,.;/]+/"
Dim sHost,sPath
'http://localhost/get/sample.asp
Set sMatch=re.Execute(sUrl)
'http://localhost
sHost=sMatch(0).SubMatches(0)
'http://localhost/get/
sPath=sMatch(0)
re.Pattern = "(src|href)=""?((?!http://)[-A-Z0-9+&@#%=~_|!:,.;/]+)""?"
Set RemoteFile = re.Execute(sContent)
'RemoteFile ������ʽMatch����ļ���
'RemoteFileUrl ������ʽMatch����,����src="Upload/a.jpg"
Dim sAbsoluteUrl
For Each RemoteFileUrl in RemoteFile
'<img src="a.jpg">,<img src="f/a.jpg">,<img src="/ff/a.jpg">
If Left(RemoteFileUrl.SubMatches(1),1)="/" Then
sAbsoluteUrl=sHost
Else
sAbsoluteUrl=sPath
End If
sAbsoluteUrl = RemoteFileUrl.SubMatches(0)&"="""&sAbsoluteUrl&RemoteFileUrl.SubMatches(1)&""""
sContent=Replace(sContent,RemoteFileUrl,sAbsoluteUrl)
Next
DetectUrl=sContent
End Function

function str2q(str)
    rem �������ַ�Ϊȫ��
    dim dist
    dist=replace(str,"'","��")
    dist=replace(dist,chr(34),"��")
    dist=replace(dist,"<","��")
    dist=replace(dist,">","��")
    dist=replace(dist,"(","��")
    dist=replace(dist,")","��")
    ''dist=replace(dist,";;","����")
    str2q=dist
end function

function str2b(str)
    rem �������ַ�Ϊ���
    dim dist
    dist=replace(str,"��","'")
    dist=replace(dist,"��",chr(34))
    dist=replace(dist,"��","<")
    dist=replace(dist,"��",">")
    dist=replace(dist,"��","(")
    dist=replace(dist,"��",")")
    ''dist=replace(dist,"����",";;")
    str2b=dist
end function

'==================================================
'��������delalert
'��  �ã�ɾ������ȷ�ϡ�ȡ���Ի���
'��  ����strurl--������Ϣ����ַ��������ʾ���֣��á�|��������
'��  �ӣ�delalert "�Ƿ����C9CMS|http://c9cms.cn|����Ϥ��"
'==================================================
Function delalert(strurl)
 dim str,mgs,del,a,b,c,d
  str=split(strurl,"|")
   mgs="��ܰ��ʾ��\r\n��ȷ��Ҫɾ����\r\nɾ����ɲ��ָܻ���Ŷ��"
    del="ɾ��"
    a="<a href=""#"" onclick='if(!confirm("""
   b=""")){return false;}else{window.document.location.href="""
  c=""";}'>"
 d="</a>"
  If str(0)="" and str(1)<>"" and str(2)<>"" then
    Wstr a&mgs&b&str(1)&c&str(2)&d
    Exit Function
    Elseif str(0)="" and str(1)<>"" and str(2)="" then
    Wstr a&mgs&b&str(1)&c&del&d
    Exit Function
    Elseif str(0)<>"" and str(1)<>"" and str(2)="" then
    Wstr a&str(0)&b&str(1)&c&del&d
    Exit Function
    Elseif str(0)<>"" and str(1)<>"" and str(2)<>"" then
    Wstr a&str(0)&b&str(1)&c&str(2)&d
    Exit Function
    Else
    Wstr a&mgs&b&str(1)&c&del&d
    Exit Function
  End If
End Function

Sub Wstr(str)
    Response.Write str
End Sub

Sub CrzColor(color,str)
    If color="" then
      Wstr "<Font color=""Red"">"&str&"</Font>"
      Else
       Wstr "<Font color="""&color&""">"&str&"</Font>"
    End If
End Sub

'=================================================
'���̣� Sleep
'���ܣ� �����ڴ˕�ͣ����
'������ iSeconds    Ҫ��ͣ������
'˵���� �ú�����ʱ��δ���ϣ��������¸��汾���õ���!
'=================================================
Sub Sleep(iSeconds)
    Wstr "<br><font color=blue>��ʼ��ͣ "&iSeconds&" ��</font><br>"
    Dim t:t=Timer()
    While(Timer()<t+iSeconds)
    Response.Flush '��ʱ���
    Response.Clear
        'Do Nothing
    Wend
    Wstr "<font color=blue>��ͣ "&iSeconds&" �����</font><br>"
End Sub

%>