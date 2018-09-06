<?php
include "excel.class.php";

$excel = new Excel();
$excel->setColumns('id','name','email');

$excel->addRow('1','Commodo Cursus','cursus@xyz.com');
$excel->addRow('2','Condimentum Egestas Vulputate','condimentum@xyz.com');
$excel->addRow('3','Fusce Ultricies','fusce@xyz.com');
$excel->addRow('4','Bibendum Vulputate','bibendum@xyz.com');
$excel->addRow('5','Amet Risus Nullam','amet@xyz.com');

$excel->addRows(
    array('6','Ultricies Elit','ultricies@xyz.com'),
    array('7','Lorem Commodo Etiam','lorem@xyz.com'),
    array('8','Vulputate','vulputate@xyz.com'),
    array('9','Condimentum Sit Commodo Pharetra','condimentum@xyz.com'),
    array('10','Vulputate Fusce','vulputate@xyz.com')
);

//$excel->save();
$excel->download();