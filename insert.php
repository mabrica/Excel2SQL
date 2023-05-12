<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// Excelブックの読み込み
try {
	$fileName = isset($argv[1]) ? $argv[1] : 'insert.xlsx';
	$fileType = IOFactory::identify($fileName);
	$reader = IOFactory::createReader($fileType);
	$book = $reader->load($fileName);
} catch (PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
    echo $e;
	exit;
}


// クエリ文を配列で格納
$queries = [];

// シート数を取得
$sheetCount = $book->getSheetCount();

for ($i = 0; $i < $sheetCount; $i++) {

	// シート
	$sheet = $book->getSheet($i);

	// シート名
	$sheetName = $sheet->getTitle();
	
	// クエリ文
	$queries[$sheetName] = "insert into {$sheetName} ";

	foreach ($sheet->getRowIterator() as $row) {

		$r = $row->getRowIndex();

		$queries[$sheetName] .= '(';

		foreach($sheet->getColumnIterator() as $column)
		{
			$c = $column->getColumnIndex();

			$cell = $sheet->getCell($c . $r);
			$value = $cell->getValue();
			$datatype = $cell->getDataType();

			if ($r > 1 && $datatype == 'null') {
				$queries[$sheetName] .= 'null';
			} elseif ($r > 1 && $datatype == 's') {
				$queries[$sheetName] .= "'{$value}'";
			} else {
				$queries[$sheetName] .= $value;
			}

			if ($c != $sheet->getHighestColumn()) {
				$queries[$sheetName] .= ', ';
			}

		}

		$queries[$sheetName] .= ')';

		if ($r == 1) {
			$queries[$sheetName] .= ' values ';
		} elseif ($r == $sheet->getHighestRow()) {
			$queries[$sheetName] .= ';';
		} else {
			$queries[$sheetName] .= ', ';
		}

	}
}

// ファイルを出力
if (!file_exists('sqls')) {
	mkdir('sqls', 0777);
}
foreach ($queries as $sheetName => $query) {
	file_put_contents("./sqls/{$sheetName}.sql", $query);
}