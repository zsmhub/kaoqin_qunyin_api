# 群英考勤机对接EHR系统
- 将群英考勤机的打卡记录导入EHR系统并实现群英考勤机的相关接口api

# 群英考勤机接口说明,请看类Attendance的公有函数
1. attmaster -> 获取考勤机的打卡记录并存入EHR表ATTMASTER中
2. get_department -> 获取公司部门
3. get_employee_all -> 导出全部公司员工
4. add_employee_batch -> 读取excel向群英考勤机批量导入员工
5. get_device -> 获取接入设备列表

# 项目正式上线教程：
1. 修改./index.php文件以下变量：
- $effective_date  //项目上线日期
- $api  //群英云考勤接口API
- $secrect   //群英云考勤接口通讯密钥
- $serverName  //EHR数据库访问数据库ip
- $database  //EHR数据库访问数据库名
- $uid  //EHR数据库访问数据库登陆账号
- $pwd  //EHR数据库访问数据库登陆密码
2. 在./config/EHR_EMPLOYEE.xlsx补充员工
- EHR、中控考勤机员工卡号和群英考勤机员工卡号关联映射表
- 在EHR系统上，旧员工使用中控考勤机的卡号，新员工使用群英考勤机的卡号
