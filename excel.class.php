<?php
/**
 * Created by: Carlos Henrique Fagundes
 * https://github.com/kiaonline
 * 
 *  A simple way to create a xls file 
 * 
 */
class Excel{
    private $data = array();

    function __construct(Array ...$columns){
        if(!empty($columns)) $this->setColumns($columns);
    }
    
    function getData(){
        return $this->data;
    }

    function setData(Array $data){
        $this->data = $data;
    }

    function setColumns(...$columns){
        if(!empty($columns)){
            array_unshift($this->data, $columns);
        }
    }

    function addRow(...$data){
        $this->data[] = $data;
    }
    function addRows(Array ...$data){
        foreach($data as $item){
            array_push($this->data,$item);
        }
    }

    private function BOF() {
        return pack("ssssss", 0x809, 0x8, 0x0, 0x10, 0x0, 0x0);
    }

    private function EOF() {
        return pack("ss", 0x0A, 0x00);
    }

    private function colNumber($r, $c, $v) {
        $str =  pack("sssss", 0x203, 14, $r, $c, 0x0);
        $str .= pack("d", $v);
        return $str;
    }

    private function colString($r, $c, $v) {
        $l = strlen($v);
        $str = pack("ssssss", 0x204, 8 + $l, $r, $c, 0x0, $l);
        $str .= $v;
        return $str;
    } 
    public function toJson(){
		return json_encode($this->data);
    }
    private function processData(){
        $str = $this->BOF();
        foreach($this->data as $row_index => $row){

            foreach($row as $col_index => $value){
                $isNum  = ctype_digit($value);
                if($isNum){
                    $str .= $this->colNumber($row_index, $col_index, utf8_decode($value));
                }else{
                    $str .= $this->colString($row_index, $col_index, utf8_decode($value));
                }
            }

        }
        $str .= $this->EOF();
        return $str;
    }
    function download($filename = null){
        
        if(!$filename) $filename = date("Y-m-d");
        $filename .= ".xls";

        header("Content-Type: application/force-download");
        header("Content-Type: application/octet-stream");
        header("Content-Type: application/download");
        header("Content-Disposition: attachment; filename=\"" . $filename . "\"");
        header("Content-Transfer-Encoding: binary");
        header("Pragma: no-cache");
        header("Expires: 0");
        echo $this->processData();
    }
    function save($filename = null,$path = "./"){
        if(!$filename) $filename = date("Y-m-d");
        $filename .= ".xls";
        $contents = $this->processData();
        return file_put_contents("{$path}{$filename}", $contents);
    }
}
