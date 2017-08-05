<%
'*********************
'<><>< 返回中文拼音 ><><>
'*********************

Class pingyinClass
  
  public pyDict
  
  
  '***********************************
  '   类初始化
  '***********************************
  private sub Class_Initialize
  
	  Set d = CreateObject("Scripting.Dictionary") 
		  d.add "A",-20319 
		  d.add "Ai",-20317 
		  d.add "An",-20304 
		  d.add "Ang",-20295 
		  d.add "Ao",-20292 
		  d.add "Ba",-20283 
		  d.add "Bai",-20265 
		  d.add "Ban",-20257 
		  d.add "Bang",-20242 
		  d.add "Bao",-20230 
		  d.add "Bei",-20051 
		  d.add "Ben",-20036 
		  d.add "Beng",-20032 
		  d.add "Bi",-20026 
		  d.add "Bian",-20002 
		  d.add "Biao",-19990 
		  d.add "Bie",-19986 
		  d.add "Bin",-19982 
		  d.add "Bing",-19976 
		  d.add "Bo",-19805 
		  d.add "Bu",-19784 
		  d.add "Ca",-19775 
		  d.add "Cai",-19774 
		  d.add "Can",-19763 
		  d.add "Cang",-19756 
		  d.add "Cao",-19751 
		  d.add "Ce",-19746 
		  d.add "Ceng",-19741 
		  d.add "Cha",-19739 
		  d.add "Chai",-19728 
		  d.add "Chan",-19725 
		  d.add "Chang",-19715 
		  d.add "Chao",-19540 
		  d.add "Che",-19531 
		  d.add "Chen",-19525 
		  d.add "Cheng",-19515 
		  d.add "Chi",-19500 
		  d.add "Chong",-19484 
		  d.add "Chou",-19479 
		  d.add "Chu",-19467 
		  d.add "Chuai",-19289 
		  d.add "Chuan",-19288 
		  d.add "Chuang",-19281 
		  d.add "Chui",-19275 
		  d.add "Chun",-19270 
		  d.add "Chuo",-19263 
		  d.add "Ci",-19261 
		  d.add "Cong",-19249 
		  d.add "Cou",-19243 
		  d.add "Cu",-19242 
		  d.add "Cuan",-19238 
		  d.add "Cui",-19235 
		  d.add "Cun",-19227 
		  d.add "Cuo",-19224 
		  d.add "Da",-19218 
		  d.add "Dai",-19212 
		  d.add "Dan",-19038 
		  d.add "Dang",-19023 
		  d.add "Dao",-19018 
		  d.add "De",-19006 
		  d.add "Deng",-19003 
		  d.add "Di",-18996 
		  d.add "Dian",-18977 
		  d.add "Diao",-18961 
		  d.add "Die",-18952 
		  d.add "Ding",-18783 
		  d.add "Diu",-18774 
		  d.add "Dong",-18773 
		  d.add "Dou",-18763 
		  d.add "Du",-18756 
		  d.add "Duan",-18741 
		  d.add "Dui",-18735 
		  d.add "Dun",-18731 
		  d.add "Duo",-18722 
		  d.add "E",-18710 
		  d.add "En",-18697 
		  d.add "Er",-18696 
		  d.add "Fa",-18526 
		  d.add "Fan",-18518 
		  d.add "Fang",-18501 
		  d.add "Fei",-18490 
		  d.add "Fen",-18478 
		  d.add "Feng",-18463 
		  d.add "Fo",-18448 
		  d.add "Fou",-18447 
		  d.add "Fu",-18446 
		  d.add "Ga",-18239 
		  d.add "Gai",-18237 
		  d.add "Gan",-18231 
		  d.add "Gang",-18220 
		  d.add "Gao",-18211 
		  d.add "Ge",-18201 
		  d.add "Gei",-18184 
		  d.add "Gen",-18183 
		  d.add "Geng",-18181 
		  d.add "Gong",-18012 
		  d.add "Gou",-17997 
		  d.add "Gu",-17988 
		  d.add "Gua",-17970 
		  d.add "Guai",-17964 
		  d.add "Guan",-17961 
		  d.add "Guang",-17950 
		  d.add "Gui",-17947 
		  d.add "Gun",-17931 
		  d.add "Guo",-17928 
		  d.add "Ha",-17922 
		  d.add "Hai",-17759 
		  d.add "Han",-17752 
		  d.add "Hang",-17733 
		  d.add "Hao",-17730 
		  d.add "He",-17721 
		  d.add "Hei",-17703 
		  d.add "Hen",-17701 
		  d.add "Heng",-17697 
		  d.add "Hong",-17692 
		  d.add "Hou",-17683 
		  d.add "Hu",-17676 
		  d.add "Hua",-17496 
		  d.add "Huai",-17487 
		  d.add "Huan",-17482 
		  d.add "Huang",-17468 
		  d.add "Hui",-17454 
		  d.add "Hun",-17433 
		  d.add "Huo",-17427 
		  d.add "Ji",-17417 
		  d.add "Jia",-17202 
		  d.add "Jian",-17185 
		  d.add "Jiang",-16983 
		  d.add "Jiao",-16970 
		  d.add "Jie",-16942 
		  d.add "Jin",-16915 
		  d.add "Jing",-16733 
		  d.add "Jiong",-16708 
		  d.add "Jiu",-16706 
		  d.add "Ju",-16689 
		  d.add "Juan",-16664 
		  d.add "Jue",-16657 
		  d.add "Jun",-16647 
		  d.add "Ka",-16474 
		  d.add "Kai",-16470 
		  d.add "Kan",-16465 
		  d.add "Kang",-16459 
		  d.add "Kao",-16452 
		  d.add "Ke",-16448 
		  d.add "Ken",-16433 
		  d.add "Keng",-16429 
		  d.add "Kong",-16427 
		  d.add "Kou",-16423 
		  d.add "Ku",-16419 
		  d.add "Kua",-16412 
		  d.add "Kuai",-16407 
		  d.add "Kuan",-16403 
		  d.add "Kuang",-16401 
		  d.add "Kui",-16393 
		  d.add "Kun",-16220 
		  d.add "Kuo",-16216 
		  d.add "La",-16212 
		  d.add "Lai",-16205 
		  d.add "Lan",-16202 
		  d.add "Lang",-16187 
		  d.add "Lao",-16180 
		  d.add "Le",-16171 
		  d.add "Lei",-16169 
		  d.add "Leng",-16158 
		  d.add "Li",-16155 
		  d.add "Lia",-15959 
		  d.add "Lian",-15958 
		  d.add "Liang",-15944 
		  d.add "Liao",-15933 
		  d.add "Lie",-15920 
		  d.add "Lin",-15915 
		  d.add "Ling",-15903 
		  d.add "Liu",-15889 
		  d.add "Long",-15878 
		  d.add "Lou",-15707 
		  d.add "Lu",-15701 
		  d.add "Lv",-15681 
		  d.add "Luan",-15667 
		  d.add "Lue",-15661 
		  d.add "Lun",-15659 
		  d.add "Luo",-15652 
		  d.add "Ma",-15640 
		  d.add "Mai",-15631 
		  d.add "Man",-15625 
		  d.add "Mang",-15454 
		  d.add "Mao",-15448 
		  d.add "Me",-15436 
		  d.add "Mei",-15435 
		  d.add "Men",-15419 
		  d.add "Meng",-15416 
		  d.add "Mi",-15408 
		  d.add "Mian",-15394 
		  d.add "Miao",-15385 
		  d.add "Mie",-15377 
		  d.add "Min",-15375 
		  d.add "Ming",-15369 
		  d.add "Miu",-15363 
		  d.add "Mo",-15362 
		  d.add "Mou",-15183 
		  d.add "Mu",-15180 
		  d.add "Na",-15165 
		  d.add "Nai",-15158 
		  d.add "Nan",-15153 
		  d.add "Nang",-15150 
		  d.add "Nao",-15149 
		  d.add "Ne",-15144 
		  d.add "Nei",-15143 
		  d.add "Nen",-15141 
		  d.add "Neng",-15140 
		  d.add "Ni",-15139 
		  d.add "Nian",-15128 
		  d.add "Niang",-15121 
		  d.add "Niao",-15119 
		  d.add "Nie",-15117 
		  d.add "Nin",-15110 
		  d.add "Ning",-15109 
		  d.add "Niu",-14941 
		  d.add "Nong",-14937 
		  d.add "Nu",-14933 
		  d.add "Nv",-14930 
		  d.add "Nuan",-14929 
		  d.add "Nue",-14928 
		  d.add "Nuo",-14926 
		  d.add "O",-14922 
		  d.add "Ou",-14921 
		  d.add "Pa",-14914 
		  d.add "Pai",-14908 
		  d.add "Pan",-14902 
		  d.add "Pang",-14894 
		  d.add "Pao",-14889 
		  d.add "Pei",-14882 
		  d.add "Pen",-14873 
		  d.add "Peng",-14871 
		  d.add "Pi",-14857 
		  d.add "Pian",-14678 
		  d.add "Piao",-14674 
		  d.add "Pie",-14670 
		  d.add "Pin",-14668 
		  d.add "Ping",-14663 
		  d.add "Po",-14654 
		  d.add "Pu",-14645 
		  d.add "Qi",-14630 
		  d.add "Qia",-14594 
		  d.add "Qian",-14429 
		  d.add "Qiang",-14407 
		  d.add "Qiao",-14399 
		  d.add "Qie",-14384 
		  d.add "Qin",-14379 
		  d.add "Qing",-14368 
		  d.add "Qiong",-14355 
		  d.add "Qiu",-14353 
		  d.add "Qu",-14345 
		  d.add "Quan",-14170 
		  d.add "Que",-14159 
		  d.add "Qun",-14151 
		  d.add "Ran",-14149 
		  d.add "Rang",-14145 
		  d.add "Rao",-14140 
		  d.add "Re",-14137 
		  d.add "Ren",-14135 
		  d.add "Reng",-14125 
		  d.add "Ri",-14123 
		  d.add "Rong",-14122 
		  d.add "Rou",-14112 
		  d.add "Ru",-14109 
		  d.add "Ruan",-14099 
		  d.add "Rui",-14097 
		  d.add "Run",-14094 
		  d.add "Ruo",-14092 
		  d.add "Sa",-14090 
		  d.add "Sai",-14087 
		  d.add "San",-14083 
		  d.add "Sang",-13917 
		  d.add "Sao",-13914 
		  d.add "Se",-13910 
		  d.add "Sen",-13907 
		  d.add "Seng",-13906 
		  d.add "Sha",-13905 
		  d.add "Shai",-13896 
		  d.add "Shan",-13894 
		  d.add "Shang",-13878 
		  d.add "Shao",-13870 
		  d.add "She",-13859 
		  d.add "Shen",-13847 
		  d.add "Sheng",-13831 
		  d.add "Shi",-13658 
		  d.add "Shou",-13611 
		  d.add "Shu",-13601 
		  d.add "Shua",-13406 
		  d.add "Shuai",-13404 
		  d.add "Shuan",-13400 
		  d.add "Shuang",-13398 
		  d.add "Shui",-13395 
		  d.add "Shun",-13391 
		  d.add "Shuo",-13387 
		  d.add "Si",-13383 
		  d.add "Song",-13367 
		  d.add "Sou",-13359 
		  d.add "Su",-13356 
		  d.add "Suan",-13343 
		  d.add "Sui",-13340 
		  d.add "Sun",-13329 
		  d.add "Suo",-13326 
		  d.add "Ta",-13318 
		  d.add "Tai",-13147 
		  d.add "Tan",-13138 
		  d.add "Tang",-13120 
		  d.add "Tao",-13107 
		  d.add "Te",-13096 
		  d.add "Teng",-13095 
		  d.add "Ti",-13091 
		  d.add "Tian",-13076 
		  d.add "Tiao",-13068 
		  d.add "Tie",-13063 
		  d.add "Ting",-13060 
		  d.add "Tong",-12888 
		  d.add "Tou",-12875 
		  d.add "Tu",-12871 
		  d.add "Tuan",-12860 
		  d.add "Tui",-12858 
		  d.add "Tun",-12852 
		  d.add "Tuo",-12849 
		  d.add "Wa",-12838 
		  d.add "Wai",-12831 
		  d.add "Wan",-12829 
		  d.add "Wang",-12812 
		  d.add "Wei",-12802 
		  d.add "Wen",-12607 
		  d.add "Weng",-12597 
		  d.add "Wo",-12594 
		  d.add "Wu",-12585 
		  d.add "Xi",-12556 
		  d.add "Xia",-12359 
		  d.add "Xian",-12346 
		  d.add "Xiang",-12320 
		  d.add "Xiao",-12300 
		  d.add "Xie",-12120 
		  d.add "Xin",-12099 
		  d.add "Xing",-12089 
		  d.add "Xiong",-12074 
		  d.add "Xiu",-12067 
		  d.add "Xu",-12058 
		  d.add "Xuan",-12039 
		  d.add "Xue",-11867 
		  d.add "Xun",-11861 
		  d.add "Ya",-11847 
		  d.add "Yan",-11831 
		  d.add "Yang",-11798 
		  d.add "Yao",-11781 
		  d.add "Ye",-11604 
		  d.add "Yi",-11589 
		  d.add "Yin",-11536 
		  d.add "Ying",-11358 
		  d.add "Yo",-11340 
		  d.add "Yong",-11339 
		  d.add "You",-11324 
		  d.add "Yu",-11303 
		  d.add "Yuan",-11097 
		  d.add "Yue",-11077 
		  d.add "Yun",-11067 
		  d.add "Za",-11055 
		  d.add "Zai",-11052 
		  d.add "Zan",-11045 
		  d.add "Zang",-11041 
		  d.add "Zao",-11038 
		  d.add "Ze",-11024 
		  d.add "Zei",-11020 
		  d.add "Zen",-11019 
		  d.add "Zeng",-11018 
		  d.add "Zha",-11014 
		  d.add "Zhai",-10838 
		  d.add "Zhan",-10832 
		  d.add "Zhang",-10815 
		  d.add "Zhao",-10800 
		  d.add "Zhe",-10790 
		  d.add "Zhen",-10780 
		  d.add "Zheng",-10764 
		  d.add "Zhi",-10587 
		  d.add "Zhong",-10544 
		  d.add "Zhou",-10533 
		  d.add "Zhu",-10519 
		  d.add "Zhua",-10331 
		  d.add "Zhuai",-10329 
		  d.add "Zhuan",-10328 
		  d.add "Zhuang",-10322 
		  d.add "Zhui",-10315 
		  d.add "Zhun",-10309 
		  d.add "Zhuo",-10307 
		  d.add "Zi",-10296 
		  d.add "Zong",-10281 
		  d.add "Zou",-10274 
		  d.add "Zu",-10270 
		  d.add "Zuan",-10262 
		  d.add "Zui",-10260
		  d.add "Zun",-10256
		  d.add "Zuo",-10254
		  set pyDict = d
	  Set d = nothing
	  
  End Sub
  
  
  public function pingyinCode(num) 
	  pingyinCode = ""
	  if (num>47 and num<58) or (num>64 and num<91) or (num>96 and num<123) then 
		 pingyinCode = chr(num) 
	  elseif num >= -20319 and num <= -10247 then 
		  a = pyDict.Items 
		  b = pyDict.keys 
		  for i = pyDict.count-1 to 0 step -1 
			  if a(i)<=num then exit for 
		  next 
		  pingyinCode = b(i)
	  else
		  pingyinCode = "_"
	  end if
  end function
  
  public function show(str) 
	  show = ""
	  for p_i=1 to len(str)
		  show = show & pingyinCode( asc(mid(str,p_i,1)) )
	  next
  end function
 
End Class

  
'*********************
'<><>< 实例化对象 ><><>
'*********************
set pingyin = New pingyinClass

response.Write(pingyin.show("中国"))
%>