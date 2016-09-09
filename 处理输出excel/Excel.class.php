<?php
/**
 * 生成excel文件操作
 *
 * @author wesley wu
 * @date 2013.12.9
 */
class Excel
{
     
    private $limit = 10000;
     
    public function download($data, $fileName)
    {
        $fileName = $this->_charset($fileName);
        header("Content-Type: application/vnd.ms-excel; charset=utf8");
        header("Content-Disposition: inline; filename=\"" . $fileName . ".xls\"");
        echo "<?xml version=\"1.0\" encoding=\"utf8\"?>\n
            <Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"
            xmlns:x=\"urn:schemas-microsoft-com:office:excel\"
            xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"
            xmlns:html=\"http://www.w3.org/TR/REC-html40\"><Styles>
            <Style ss:ID=\"s69\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/>
   <Font ss:FontName=\"宋体\" x:CharSet=\"134\" ss:Size=\"24\" ss:Color=\"#000000\"/></Style></Styles>";
        echo "\n<Worksheet ss:Name=\"" . $fileName . "\">\n<Table>\n";
        $guard = 0;
        $ii=0;
        foreach($data as $v)
        {
            $ii++;
            $guard++;
            if($guard==$this->limit)
            {
                ob_flush();
                flush();
                $guard = 0;
            }
            echo $this->_addRow($this->_charset($v),$ii);
        }
        echo "</Table>\n</Worksheet>\n</Workbook>";
    }

    private function _addRow($row,$oo)
    {
        $cells = "";
        $i=0;
        foreach ($row as $k => $v)
        {
            $i++;
            if($i==1 && $oo==1){
              $cells.="<Cell ss:MergeAcross=\"17\" ss:StyleID=\"s69\"><Data ss:Type=\"String\">".$v."</Data></Cell>\n";
            }else{
                if($v==''){
                    $v=0;
                }
                $cells.="<Cell><Data ss:Type=\"String\">".$v."</Data></Cell>\n";
            }
        }
        if($oo==1){
            return "<Row ss:AutoFitHeight=\"0\" ss:Height=\"46\">\n" . $cells . "</Row>\n";
        }else{
            return "<Row>\n" . $cells . "</Row>\n";
        }

    }
     
    private function _charset($data)
    {
        if(!$data)
        {
            return false;
        }
        if(is_array($data))
        {
            foreach($data as $k=>$v)
            {
                $data[$k] = $this->_charset($v);
            }
            return $data;
        }
        return iconv('utf-8', 'utf-8', $data);
    }
     
}

/*$excel = new Excel();
$data = array(
    array('姓名','标题','文章','价格','数据5','数据6','数据7'),
    array('数据1','数据2','数据3','数据4','数据5','数据6','数据7'),
    array('数据1','数据2','数据3','数据4','数据5','数据6','数据7'),
    array('数据1','数据2','数据3','数据4','数据5','数据6','数据7'),
    array('数据1','数据2','数据3','数据4','数据5','数据6','数据7'),
    array('数据1','数据2','数据3','数据4','数据5','数据6','数据7')
);
$excel->download($data, '这是一个测试');*/
?>