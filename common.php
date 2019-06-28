<?php

function curl_post_https($url, $param) {
    $str = json_encode($param);
    $ch = curl_init();
    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_POST, true);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true); //获取的信息以文件流的形式返回
    curl_setopt($ch, CURLOPT_POSTFIELDS, $str);
    curl_setopt($ch, CURLOPT_HTTPHEADER, array('Content-type:application/json'));
    $data = curl_exec($ch); //运行curl
    //halt(curl_errno($ch));
    curl_close($ch);
    return $data; //返回json对象
}

$config['token'] = 'e0cc21c2-8ddc-45bb-bbe0-ea0e3a6b5d3d';
$config['fundamental'] = 'https://open.lixinger.com/api/a/stock/fundamental';
$config['fs'] = 'https://open.lixinger.com/api/a/stock/fs/industry';
?>
