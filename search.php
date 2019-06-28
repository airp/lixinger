<html>
<head>
    <!-- 解决弹窗对话框乱码问题 -->
    <meta http-equiv="Content-Type" content="text/html;charset=utf-8">
</head>
 
<?php

if(isset($_POST["Submit"]) && $_POST["Submit"] == "download"){
    $time_info = $_POST["time_info"];
    $textarea_info = $_POST["textarea_info"];

    $week_idx = date("w",strtotime($time_info));

    if($time_info == "" || $textarea_info == "")
    {
        echo "<script>alert('请确认信息完整性！'); history.go(-1);</script>";
    }
    else if($week_idx == '0' || $week_idx == '6')
    {
        echo "<script>alert('请确认日期是否为交易日！'); history.go(-1);</script>";
    }
    else
    {
        //检测代码格式
        $property = trim($textarea_info);
        $propertyArr = explode("\n", str_replace("\r\n", "\n", $property));
        $flag = false;
        foreach ($propertyArr as $k => &$v){
            $v = trim($v);

            if(empty($v))
                unset($propertyArr[$k]);
            else
            {
                //只取 000848.SZ 的前半部分
                $tmp_v = explode('.', $v);
                $v = $tmp_v[0];

                //必须要保证长度为6，且必须是数字
                if(strlen($v) != 6 || !is_numeric($v))
                {
                    $flag = true;
                    break;
                }
            }
        }

        if($flag)
        {
            echo "<script>alert('请确认股票代码是否正确！'); history.go(-1);</script>";
        }
        else
        {
            require './all.php';
            $cfg['date'] = $time_info;
            $cfg['code'] = array_values($propertyArr);//这一步是重新建立索引

            //开始导出
            export($cfg);
        }
    }
}else{
    echo "<script>alert('提交未成功！'); history.go(-1);</script>";
}
?>
