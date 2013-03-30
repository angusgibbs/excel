<?php

require('../lib/excel.php');

$book = new Excel;
$book->sheet('Sheet1', function($sheet) {
	$sheet->populate(array(
		array('Name', 'Age'),
		array('Person', 30),
		array('Joe', 25),
		array('Bob', 65)
	));
});

echo '<pre>' . htmlspecialchars($book->end()) . '</pre>';