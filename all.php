<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('PRC');

/** 引入PHPExcel */
require_once dirname(__FILE__) . '/PHPExcel/Classes/PHPExcel.php';
require './common.php';

function export($cfg)
{
    global $config;

    $url = $config['fundamental'];
    $str = array(
        'token' => $config['token'],
        'date' => $cfg['date'],
        'stockCodes' => $cfg['code'],
        'metrics' => ['d_pe_ttm_pos10','pb_wo_gw_pos10','dyr','sp'],
    );

    //查询PE & PB
    $result=curl_post_https($url, $str);
    $data=json_decode($result, true);
    $arr = $data['data'];

    $url_fs = $config['fs'];

    $offset = 4;        //连续年数
    $end_year = date("Y", strtotime($cfg['date']));
    $start_year = $end_year - $offset;

    $startDate = $start_year.'-01-01';
    $endDate = $end_year.'-01-01';

    $param = array(
        'token' => $config['token'],
        'startDate' => $startDate,
        'endDate' => $endDate,
        'metrics' => ['q.profitStatement.oi.t','q.balanceSheet.ar.t', 'q.balanceSheet.i.t', 'q.balanceSheet.tca_tcl_r.t'],
    );

    $ret_data = array();
    foreach($arr as $obj)
    {
        $info['code'] = $obj['stockCode'];
        $info['d_pe'] = sprintf("%01.2f", $obj['d_pe_ttm_pos10']*100).'%';
        $info['d_pb'] = sprintf("%01.2f", $obj['pb_wo_gw_pos10']*100).'%';
        $info['dyr'] = sprintf("%01.2f", $obj['dyr']*100).'%';
        $info['sp'] = $obj['sp'];

        //查询单个信息
        $param['stockCodes'] = array(
            $obj['stockCode'],
        );

        $res=curl_post_https($url_fs, $param);
        $res_data=json_decode($res, true);
        $arr_fs = $res_data['data'];

        $info['year_list'] = array();
        foreach($arr_fs as $ob)
        {
            if($ob['reportType'] == 'annual_report')
            {
                $year = date("Y", strtotime($ob['date']));
                $oi = $ob['q']['profitStatement']['oi']['t'];
                $balanceSheet=$ob['q']['balanceSheet'];
                $ar = isset($balanceSheet['ar']) ? $balanceSheet['ar']['t'] : 0;
                $i = isset($balanceSheet['i']) ? $balanceSheet['i']['t'] : 0;
                $tca_tcl_r = isset($balanceSheet['tca_tcl_r']) ? $balanceSheet['tca_tcl_r']['t'] : 0;
                $oi = sprintf("%01.2f", $oi / 100000000);
                $ar = sprintf("%01.2f", $ar / 100000000);
                $i = sprintf("%01.2f", $i / 100000000);
                $tca_tcl_r = sprintf("%01.2f", $tca_tcl_r);

                $year_info['year'] = $year;
                $year_info['oi'] = $oi;
                $year_info['ar'] = $ar;
                $year_info['i'] = $i;
                $year_info['tca_tcl_r'] = $tca_tcl_r;
                array_push($info['year_list'], $year_info);
            }
        }

        array_push($ret_data, $info);
    }
    //echo json_encode($ret_data);

    //Excel列字典
    $letter = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');

    $cols = 5;          //前5列是CODE PE PB 股息率 股价

    //各项指标, 初始位置
    $yysr_pos = $cols;
    $yszk_pos = $cols + $offset * 1;
    $ch_pos = $cols + $offset * 2;
    $ldbl_pos = $cols + $offset * 3;

    // 创建Excel导出数据
    $objPHPExcel = new PHPExcel();
    // 设置文档信息，这个文档信息windows系统可以右键文件属性查看
    $objPHPExcel->getProperties()->setCreator("Airp")
                ->setLastModifiedBy("Airp")
                ->setTitle("lixinger")
                ->setSubject("lixinger")
                ->setDescription("lixinger")
                ->setKeywords("lixinger")
                ->setCategory("lixinger");

    // 设置第一个sheet为工作的sheet
    $objPHPExcel->setActiveSheetIndex(0);

    //设置列名
    $objPHPExcel->getActiveSheet()->setcellValue('A2', 'CODE');
    $objPHPExcel->getActiveSheet()->setcellValue('B2', 'PE-TTM(扣非)分位点(10年)');
    $objPHPExcel->getActiveSheet()->setcellValue('C2', 'PB(不含商誉)分位点(10年)');
    $objPHPExcel->getActiveSheet()->setcellValue('D2', '股息率');
    $objPHPExcel->getActiveSheet()->setcellValue('E2', '股价');
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
	// $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
	// $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);

    $objPHPExcel->getActiveSheet()->setcellValue($letter[$yysr_pos].'1', '营业收入');
    $objPHPExcel->getActiveSheet()->setcellValue($letter[$yszk_pos].'1', '应收账款');
    $objPHPExcel->getActiveSheet()->setcellValue($letter[$ch_pos].'1', '存货');
    $objPHPExcel->getActiveSheet()->setcellValue($letter[$ldbl_pos].'1', '流动比率');

    //合并单元格
    $objPHPExcel->getActiveSheet()->mergeCells($letter[$yysr_pos].'1:'.$letter[$yysr_pos + $offset - 1].'1');
    $objPHPExcel->getActiveSheet()->mergeCells($letter[$yszk_pos].'1:'.$letter[$yszk_pos + $offset - 1].'1');
    $objPHPExcel->getActiveSheet()->mergeCells($letter[$ch_pos].'1:'.$letter[$ch_pos + $offset - 1].'1');
    $objPHPExcel->getActiveSheet()->mergeCells($letter[$ldbl_pos].'1:'.$letter[$ldbl_pos + $offset - 1].'1');

    //设置居中
    $objPHPExcel->getActiveSheet()->getStyle($letter[$yysr_pos].'1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle($letter[$yszk_pos].'1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle($letter[$ch_pos].'1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()->getStyle($letter[$ldbl_pos].'1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

    $ix = 0;
    for($y = $end_year - 1; $y >= $start_year; --$y)
    {
        //记录年份位置
        $year_pos[$y] = $ix;

        $objPHPExcel->getActiveSheet()->setcellValue($letter[$ix + $yysr_pos].'2', $y);
        $objPHPExcel->getActiveSheet()->setcellValue($letter[$ix + $yszk_pos].'2', $y);
        $objPHPExcel->getActiveSheet()->setcellValue($letter[$ix + $ch_pos].'2', $y);
        $objPHPExcel->getActiveSheet()->setcellValue($letter[$ix + $ldbl_pos].'2', $y);
        $ix++;
    }

    $line_num = 3;      //从第三行开始
    foreach($ret_data as $obj)
    {
        $objPHPExcel->getActiveSheet()->setCellValue('A'.$line_num, $obj['code']);
        $objPHPExcel->getActiveSheet()->setCellValue('B'.$line_num, $obj['d_pe']);
        $objPHPExcel->getActiveSheet()->setCellValue('C'.$line_num, $obj['d_pb']);
        $objPHPExcel->getActiveSheet()->setCellValue('D'.$line_num, $obj['dyr']);
        $objPHPExcel->getActiveSheet()->setCellValue('E'.$line_num, $obj['sp']);

        foreach($obj['year_list'] as $ob)
        {
            //确保获取的数据, 写入Excel中正确的位置
            if(isset($year_pos[$ob['year']]))
            {
                $pos = $year_pos[$ob['year']];

                $objPHPExcel->getActiveSheet()->setCellValue($letter[$yysr_pos + $pos].$line_num, $ob['oi']);
                $objPHPExcel->getActiveSheet()->setCellValue($letter[$yszk_pos + $pos].$line_num, $ob['ar']);
                $objPHPExcel->getActiveSheet()->setCellValue($letter[$ch_pos + $pos].$line_num, $ob['i']);
                $objPHPExcel->getActiveSheet()->setCellValue($letter[$ldbl_pos + $pos].$line_num, $ob['tca_tcl_r']);
            }
        }
        $line_num++;
    }

    // 保存Excel 2007格式文件，保存路径为当前路径，名字为export.xlsx
    //$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
    //$objWriter->save('reports.xlsx');

	$filename = ('reports').'_'.date('Y-m-d_His');

    //*生成xlsx文件
	header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="'.$filename.'.xlsx"');
	header('Cache-Control: max-age=0');
	header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); 
	header('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT');
	header('Cache-Control: cache, must-revalidate'); 
	header('Pragma: public');

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	ob_end_clean();
    $objWriter->save('php://output');
    exit();
}

?>
