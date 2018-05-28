<?php
date_default_timezone_set('PRC');  //设置时区

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
* 群英考勤机对接EHR
*/
class Attendance {

    //接口正式上线日期
    private $effective_date = '2018-05-10';

    /**
    * 群英云考勤接口秘钥设置
    */
    //群英云考勤访问域名
    private $url_attendance = "http://kq.yisu.com/";
    //API帐号
    private $api = "a3f682516d932a220cfe44021qba8c71";
    //通讯密钥
    private $secrect = "secrect";

    /**
    * 数据库访问配置
    */
    //数据库ip
    // private $serverName = '113.108.232.9';
    private $serverName = '192.168.1.x';
    //数据库名
    // private $database = "FEIDONG_EHR";
    private $database = "EHR_TEST";
    //登陆账号
    // private $uid = "ehr_kaoqin";
    private $uid = "ehr-kaoqin-test";
    //登陆密码
    private $pwd = "ab@1#2";

    /**
    * 日期时间格式判断
    */
    private function is_date($date, $fmt='Y-m-d'){
        if(empty($date)) return false;
        return date($fmt,strtotime($date))== $date;
    }

    /**
     * curl请求
     */
    private function curl($url, $poststr='', $httpheader=array(), $return_header=false){
        $ch = curl_init();
        $SSL = substr($url, 0, 8) == "https://" ? true : false;
        curl_setopt($ch, CURLOPT_URL, $url);
        curl_setopt($ch, CURLOPT_HEADER, $return_header);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($ch, CURLOPT_FOLLOWLOCATION, 1);
        if ( $SSL ) {
            curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
            curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, 2);
        }
        if( $poststr!='' ){
            curl_setopt($ch, CURLOPT_POST, 1);
            curl_setopt($ch, CURLOPT_POSTFIELDS, $poststr);
        }
        if( $httpheader ){
            curl_setopt($ch, CURLOPT_HTTPHEADER, $httpheader);
        }

        $data = curl_exec($ch);
        curl_close($ch);
        return $data;
    }

    /**
    * 过滤sql与php文件操作的关键字
    * @param string $string
    * @return string
    */
    private function filter_keyword($string) {
         $keyword = 'select|insert|update|delete|truncate|\/\*|\*|\.\.\/|\.\/|union|into|load_file|outfile';
         $arr = explode('|', $keyword);
         $result = str_ireplace($arr, '', $string);
         return $result;
    }

    //创建目录
    private function create_dir($dir,$mod=0755){
        if(!is_string($dir)) return false;
        $dirarr=array();
        while(!is_dir($dir)) {
            array_unshift($dirarr,$dir);
            $dir=dirname($dir);
            $char=substr($dir, -1, 1);
            if($char=='/' || $char=='\\' || $char==':') break;
        }
        foreach($dirarr as $v) {
            if(!@mkdir($v)) return false;
            @chmod($v, $mod);
        }
        return true;
    }

    /**
    * 写文件日志
    */
    private function writelog($content, $logname=''){
        //默认使用日期作为文件名
        if(empty($logname)) $logname = date('Y-m-d', time());

        $LogDir = __DIR__ . '/logs/';
        $LogFile = $LogDir . $logname . '.log';
        if(!is_dir($LogDir)) $this->create_dir($LogDir);
        if( $fp = fopen($LogFile,'a') ){
            if (flock($fp, LOCK_EX)) { // 进行排它型锁定
                fwrite($fp, date('Y-m-d H:i:s', time()) . ' -> '. $content);
                flock($fp, LOCK_UN); // 释放锁定
            }
        }
        fclose($fp);
    }

    /**
    * 数字转换成EXCEL列标签
    */
    private function int_to_chr($index, $start = 65) {
        $str = '';
        if(floor($index / 26) > 0) {
            $str .= IntToChr(floor($index / 26)-1);
        }
        return $str . chr($index % 26 + $start);
    }

    /**
    * 读取excel文件内容
    *
    * @return array
    */
    private function read_excel($path) {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($path);
        $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);
        return $sheetData;
    }

    /**
    * 写excel文件并保存在download文件夹中
    */
    private function write_excel($data, $filename='') {
        if(empty($filename)) $filename = date('YmdHis', time());
        $dir = __DIR__ . '/download/';
        $file = $dir . $filename . '.xlsx';
        if(!is_dir($dir)) $this->create_dir($dir);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $line = 1;
        foreach($data as $row) {
            $col = 0;
            foreach($row as $r) {
                $sheet->setCellValue($this->int_to_chr($col++) . $line, $r);
            }
            $line++;
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save($file);
    }

    /**
    * 返回sqlserver执行状态
    */
    private function FormatErrors($error) {
        return array(
            'SQLSTATE' => $error[0],
            'Code' => $error[1],
            'Message' => $error[2]
        );
    }

    /**
    * 数据库sql执行
    *
    * @param str $sql 查询语句
    * @param str $type sql语句类型, 包括query, insert, update
    */
    private function sqlsrv_query($sql, $type='query') {
        //数据库ip
        $serverName = $this->serverName;
        //数据库名
        $database = $this->database;
        //登陆账号
        $uid = $this->uid;
        //登陆密码
        $pwd = $this->pwd;
        //Establishes the connection
        $conn = new PDO("sqlsrv:server = {$serverName}; Database = {$database}", $uid, $pwd);
        //Executes the query
        $get_data = $conn->query($sql);
        //Error handling
        $ret_query = $this->FormatErrors($conn->errorInfo());

        //执行状态判断
        if($ret_query['SQLSTATE'] == '00000') {
            $ret = array();
            if($type == 'query') {
                while($row = $get_data->fetch(PDO::FETCH_ASSOC)) {
                    $ret[] = $row;
                }
            }

            return array('success' => true, 'msg' => '执行成功', 'data' => $ret);
        } else {
            return array('success' => false, 'msg' => '状态码：' . $ret_query['SQLSTATE'] . ', Message:' . $ret_query['Message']);
        }
    }

    /**
    * 获取公司某段时间内的打卡记录
    *
    * @param date $start 开始日期
    * @param date $end 结束日期
    * @return array 返回打卡记录
    */
    private function recordlog($params=array()) {
        //传参
        $time = time();
        $start = (isset($params[0]) && $this->is_date($params[0])) ? $params[0] : date('Y-m-d', strtotime('-1 day', $time));
        $end = (isset($params[1]) && $this->is_date($params[1])) ? $params[1] : $start;

        //公共必传参数
        $data = array(
            'account' => $this->api, //API帐号
            'requesttime' => $time, //请求时间，与服务器时间差不能超过60秒
        );

        //接口参数
        $data['start'] = $start;
        $data['end'] = $end;

        //按key排序
        ksort($data);

        //生成签名，先转成utf-8编码再生成签名
        $sign = md5(join('', $data) . $this->secrect);
        $data['sign'] = $sign;

        $ret = $this->curl($this->url_attendance . "Api/Api/recordlog?" . http_build_query($data));
        $arr = json_decode($ret, true);

        if($arr['status'] != 1) {
            $this->writelog("考勤机的打卡记录接口调用失败：{$arr['error']}\n");
        } elseif($arr['status'] == 1 && empty($arr['data']['attendata'])) {
            $this->writelog("开始日期:{$start}, 结束日期:{$end}，这段时间没有考勤记录\n");
        } else {
            $this->writelog("考勤机的打卡记录接口调用成功，开始日期:{$start}, 结束日期:{$end}\n");
        }

        return $arr;
    }

    /**
    * 获取考勤机的打卡记录并存入EHR表ATTMASTER中
    *
    * @param date $start 开始日期
    * @param date $end 结束日期
    * @return boolean 操作结果
    */
    public function attmaster($params=array()) {
        //传参
        $time = time();
        $start = (isset($params[0]) && $this->is_date($params[0])) ? $params[0] : date('Y-m-d', strtotime('-1 day', $time));
        $end = (isset($params[1]) && $this->is_date($params[1])) ? $params[1] : $start;

        //ehr打卡记录表
        $ehr_table_attmaster = "ATTMASTER";

        //ehr考勤机卡号和EHR员工编号关联表
        $ehr_table_attcdmas = "ATTCDMAS";

        //获取打卡记录,默认获取当天的打卡记录
        $recordlog_ret = $this->recordlog($params);

        //打卡记录接口调用失败
        if($recordlog_ret['status'] != 1) return false;

        //没有考勤记录
        if(empty($recordlog_ret['data']['attendata'])) return false;

        //打卡记录
        $recordlog = $recordlog_ret['data']['attendata'];

        //读取中控考勤机和群英云考勤机切换的员工考勤编号关联映射表
        $path = __DIR__ . '/config/EHR_EMPLOYEE.xlsx';
        $excel_data = $this->read_excel($path);

        //老员工：卡号映射表数组重组索引
        $excel_data_index = array();
        foreach($excel_data as $key => $row) {
            $zk_card_id = trim($row['A']);  //中控考勤机员工卡号
            $ehr_staff_no = trim($row['B']);  //EHR员工编号
            $staff_name = trim($row['C']);  //员工姓名
            $qunyin_card_id = trim($row['D']);  //群英考勤机员工卡号

            //忽略空数据、excel标题、群英考勤机员工卡号为空的记录
            if($zk_card_id == '' || $ehr_staff_no == '' || $qunyin_card_id == '' || $key == 1 || $qunyin_card_id == NULL) continue;

            $excel_data_index[$qunyin_card_id] = array('zk_card_id' => $zk_card_id, 'ehr_staff_no' => $ehr_staff_no);
        }

        //新员工：ehr考勤机卡号和EHR员工编号关联表
        $sql = "SELECT
                -- 考勤机卡号
                CARD_ID,
                -- EHR员工编号
                STAFF_NO
                from {$ehr_table_attcdmas}
                WHERE EFFECTIVE_DATE >= '{$this->effective_date}'
                ";
        $ret_query = $this->sqlsrv_query($sql);
        $ret_attcdmas = array();  //新员工EHR员工编号表
        if(!$ret_query['success']) {
            $this->writelog("这条sql语句执行失败：{$sql}\n, 错误信息：{$ret_query['msg']}\n");
        } else {
            foreach($ret_query['data'] as $row) {
                $card_no = trim($row['CARD_ID']);
                $ret_attcdmas[$card_no] = trim($row['STAFF_NO']);
            }
        }

        //打卡记录整合
        $flag_insert = true;
        $insert_record = array();  //打卡记录
        $insert_sql = "INSERT INTO {$ehr_table_attmaster} (CARD_NO, STAFF_NO, DATES, TIME, TERMINAL_NO) values ";  //插入语句
        foreach($recordlog as $row) {
            $qunyin_card_id = $row['atten_uid'];  //群英考勤机员工卡号
            $atten_time = date('Hi', $row['atten_time']) . '00';  //打卡时间
            $atten_date = $row['atten_date'];  //打卡日期

            //设置考勤机卡号和EHR员工编号
            if(isset($excel_data_index[$qunyin_card_id])) {  //老员工使用中控考勤机的卡号
                //考勤机卡号
                $card_no = $excel_data_index[$qunyin_card_id]['zk_card_id'];

                //EHR员工编号
                $staff_no = $excel_data_index[$qunyin_card_id]['ehr_staff_no'];
            } else {  //新员工使用群英考勤机的卡号, 从EHR数据库获取数据
                $card_no = $qunyin_card_id;

                if(!isset($ret_attcdmas[$card_no])) {
                    $this->writelog("该群英考勤机卡号不存在：{$card_no}\n");
                    continue;
                }
                $staff_no = $ret_attcdmas[$card_no];
            }

            $insert_sql .= "('{$card_no}', '{$staff_no}', '{$atten_date}', '{$atten_time}', '1'),";
            $flag_insert = false;
        }
        if($flag_insert) {
            $this->writelog("没有符合要求的考勤记录\n");
            return false;
        }
        $insert_sql = substr($insert_sql, 0, -1);
        // $this->sqlsrv_query("DELETE FROM {$ehr_table_attmaster} WHERE DATES BETWEEN '{$start}' AND '{$end}'");
        $ret_insert = $this->sqlsrv_query($insert_sql);
        if(!$ret_insert['success']) {
            $this->writelog("这条sql语句执行失败：{$insert_sql}\n, 错误信息：{$ret_query['msg']}\n");
        }

        return $ret_insert;
    }

    /**
    * 获取公司部门
    */
    public function get_department() {
        //公共必传参数
        $time = time();
        $data = array(
            'account' => $this->api, //API帐号
            'requesttime' => $time, //请求时间，与服务器时间差不能超过60秒
        );

        //按key排序
        ksort($data);

        //生成签名，先转成utf-8编码再生成签名
        $sign = md5(join('', $data) . $this->secrect);
        $data['sign'] = $sign;

        $ret = $this->curl($this->url_attendance . "Api/Api/getDepartment?" . http_build_query($data));
        $arr = json_decode($ret, true);

        if($arr['status'] != 1) {
            $this->writelog("考勤机的获取公司部门接口调用失败：{$arr['error']}\n");
        } else {
            $this->writelog("考勤机的获取公司部门接口调用成功，执行日期:" . date('Y-m-d H:i:s', $time) . "\n");

            //导出excel
            $dep_regroup = array(array('部门ID', '部门名', '部门上层ID'));
            foreach($arr['data'] as $row) {
                $dep_regroup[] = array($row['id'], $row['name'], $row['pid']);
            }
            $this->write_excel($dep_regroup, 'department');
        }

        return $arr;
    }

    /**
    * 获取公司员工-按页数获取
    */
    private function get_employee($page=1) {
        //公共必传参数
        $time = time();
        $data = array(
            'account' => $this->api, //API帐号
            'requesttime' => $time, //请求时间，与服务器时间差不能超过60秒
        );

        //记录页数，每页50条记录
        $data['page'] = $page;

        //按key排序
        ksort($data);

        //生成签名，先转成utf-8编码再生成签名
        $sign = md5(join('', $data) . $this->secrect);
        $data['sign'] = $sign;

        $ret = $this->curl($this->url_attendance . "Api/Api/getEmployee?" . http_build_query($data));
        $arr = json_decode($ret, true);

        if($arr['status'] != 1) {
            $this->writelog("考勤机的获取公司员工接口调用失败：{$arr['error']}\n");
        } else {
            $this->writelog("考勤机的获取公司员工接口调用成功，执行日期:" . date('Y-m-d H:i:s', $time) . "\n");
        }

        return $arr;
    }

    /**
    * 导出全部公司员工
    */
    public function get_employee_all($params=array()) {
        $arr = $this->get_employee();
        $data = $arr['data'];
        $total = $data['total'];
        $total_page = intval($data['totalpage']);

        //第一页员工数据
        $dep_regroup = array(array('考勤编号', '姓名', '部门', '指纹数'));
        foreach($data['userData'] as $row) {
            $dep_regroup[] = array($row['account'], $row['realname'], $row['departname'], $row['fingerprint']);
        }

        //第2页及以上员工数据
        for($i = 2; $i <= $total_page; $i++) {
            $arr_next = $this->get_employee($i);
            $data_next = $arr_next['data'];
            foreach($data_next['userData'] as $row) {
                $dep_regroup[] = array($row['account'], $row['realname'], $row['departname'], $row['fingerprint']);
            }
        }

        //导出excel
        $this->write_excel($dep_regroup, 'employee');

        return array('total' => $total, 'totalpage' => $total_page);
    }

    /**
    * 单个添加员工
    */
    private function add_employee($param=array()) {
        //公共必传参数
        $time = time();
        $data = array(
            'account' => $this->api, //API帐号
            'requesttime' => $time, //请求时间，与服务器时间差不能超过60秒
        );

        $data = array_merge($data, $param);

        //按key排序
        ksort($data);

        //生成签名，先转成utf-8编码再生成签名
        $sign = md5(join('', $data) . $this->secrect);
        $data['sign'] = $sign;

        $ret = $this->curl($this->url_attendance . "Api/Api/addEmployee?" . http_build_query($data));
        $arr = json_decode($ret, true);

        if($arr['status'] != 1) {
            $this->writelog("添加员工接口调用失败：{$arr['error']}, 该员工姓名: {$data['realname']}\n");
        } else {
            $this->writelog("添加员工接口调用成功，执行日期:" . date('Y-m-d H:i:s', $time) . "\n");
        }

        return $arr;
    }

    /**
    * 读取excel批量导入员工
    * [操作流程：1. 获取./config文件夹下的群英员工批量导入模板，2. 将整理好的员工excel上传到./config文件夹下，3. 命令行调用接口批量导入员工]
    */
    public function add_employee_batch($params=array()) {
        $excel_name = trim($params[0]);
        if(empty($excel_name)) return '请选择要批量导入的员工excel名称';

        //读取中控考勤机和群英云考勤机切换的员工考勤编号关联映射表
        $path = __DIR__ . '/config/' . $excel_name . '.xlsx';
        $excel_data = $this->read_excel($path);

        $excel_data_index = array();
        foreach($excel_data as $key => $row) {
            if($key == 1) continue;

            $name = trim($row['A']);  //姓名
            $mobile = trim($row['B']);  //手机
            $department_id = trim($row['C']);  //部门ID

            //姓名、手机、部门id为必填选项
            if($name == '' || $mobile == '' || $department_id == '') {
                echo '第 ' . $key . ' 行有记录的必填选项为空，请检查' . "\n";
                continue;
            }

            //添加员工到群英考勤机系统
            $employee = array('realname' => $name, 'mobile' => $mobile, 'deptid' => $department_id);
            //群英考勤系统登录密码
            $employee['password'] = trim($row['E']) != '' ? md5(trim($row['E'])) : md5('123456');
            //性别
            if(trim($row['D']) != '') $employee['sex'] = (trim($row['D']) == '女') ? '2' : '1';
            //邮箱
            if(trim($row['F']) != '') $employee['email'] = trim($row['F']);
            //要同步设备的SN，多个用英文逗号分隔
            if(trim($row['G']) != '') $employee['sn'] = trim($row['G']);
            $this->add_employee($employee);
        }

        return true;
    }

    /**
    * 获取接入设备列表
    */
    public function get_device() {
        //公共必传参数
        $time = time();
        $data = array(
            'account' => $this->api, //API帐号
            'requesttime' => $time, //请求时间，与服务器时间差不能超过60秒
        );

        //按key排序
        ksort($data);

        //生成签名，先转成utf-8编码再生成签名
        $sign = md5(join('', $data) . $this->secrect);
        $data['sign'] = $sign;

        $ret = $this->curl($this->url_attendance . "Api/Api/getDevice?" . http_build_query($data));
        $arr = json_decode($ret, true);

        if($arr['status'] != 1) {
            $this->writelog("获取接入设备列表接口调用失败：{$arr['error']}\n");
        } else {
            $this->writelog("获取接入设备列表接口调用成功，执行日期:" . date('Y-m-d H:i:s', $time) . "\n");
        }

        return $arr;
    }
}

//访问类型，只能使用终端方式调用接口
$sapi_type = php_sapi_name();
if($sapi_type != 'cli') {
    echo '访问出错！';
    return;
}

ignore_user_abort(true);  //忽略用户断开连接
ini_set('memory_limit', -1);
set_time_limit(0);

/*
计划任务接口调用示例：
/usr/local/php/bin/php /var/www/html/kaoqin_qunyin_api/index.php attmaster 2018-04-12
*/

//接收传递参数
if($argc < 2) {
    echo "请先选择要访问的方法 \n";
    return;
}
//调用方法
$method = htmlspecialchars($argv[1]);
//调用方法对应参数
$params = $argv;
unset($params[0], $params[1]);
if(!empty($params)) sort($params);

//调用考勤接口类
$attendance = new Attendance();
if(!is_callable([$attendance, $method])) {
    echo '调用方法不存在';
    return;
}
$ret = $attendance->$method($params);

//打印结果
var_export($ret);
