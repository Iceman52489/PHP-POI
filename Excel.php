<?php

namespace Boosty\POI;

use \java AS java;

class Excel extends POI {
	private $workbook;
	private $sheets = array();
	private $cellStyle;

	private $cache = EXPORTS;
	private $filename;
	protected $saved = false;

	public $sheetNames = array();

	public function __construct($filename=NULL) {
		require_once(JAVABRIDGE.'/java.inc');

		/*----------------*/
		/*--- Workbook ---*/
		/*----------------*/
		if(isset($filename)) {
			
			$file = new java('java.io.File', $filename);
			$fileInputStream = new java('java.io.FileInputStream', $file);

			$this->workbook = new java('org.apache.poi.ss.usermodel.WorkbookFactory');
			$this->workbook = $this->workbook->create($file);

			$this->filename = $filename;

			/*--------------*/
			/*--- Sheets ---*/
			/*--------------*/
			$this->readSheets();
			$this->saved = true;
		} else {
			$this->workbook = new java('org.apache.poi.hssf.usermodel.HSSFWorkbook');
		}

		// Create CellStyle Utility Class
		$this->cellStyle = new CellStyle();

		return $this;
	}

	private function readSheets() {
		$intSheets = java_values($this->workbook->getNumberOfSheets());

		for($intSheet = 0; $intSheet < $intSheets; $intSheet++) {
			$sheetName = java_values($this->workbook->getSheetName($intSheet));

			$this->sheetNames[$intSheet] = $sheetName;
			$this->sheets[$intSheet] = new Sheet($this, $sheetName, $this->workbook->getSheetAt($intSheet));
		}
	}

	public function createSheet($sheetName=NULL) {
		$intSheet = count($this->sheets) + 1;
		$sheetName = isset($sheetName) ? $sheetName : 'Sheet - '.$intSheet;

		array_push($this->sheetNames, $sheetName);
		array_push($this->sheets, new Sheet($this, $sheetName, $this->workbook->createSheet($sheetName)));

		$this->saved = true;

		return $this->sheets[count($this->sheets) - 1];
	}

	public function getSheetAt($intSheet) {
		return $this->sheets[$intSheet];
	}

	public function getSheet($sheetName) {
		$sheetIndexes = array_flip($this->sheetNames);
		$intSheet = $sheetIndexes[$sheetName];
		return $this->getSheetAt($intSheet);
	}

	public function getSheetNames() {
		return $this->sheetNames;
	}

	public function export($filename = NULL) {
		$this->workbook->setForceFormulaRecalculation(true);

		$file = $this->write($filename);

		$fileInfo = $this->getFileInfo($file);

		$tmpFiles = glob($this->cache.'*.{pdf,xls,xlsx}');

		foreach($tmpFiles AS $tmpFile) {
			unlink($tmpFile);
		}

		unset($_COOKIE['fileDownload']);

		ob_clean();

		header('Set-Cookie: fileDownload=true; path=/');
		header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
		header('Content-Description: File Transfer');
		header('Content-Type: '.$fileInfo['mimetype']);
		header('Content-Length: '.$fileInfo['size']);
		header('Content-Disposition: attachment; filename="'.$fileInfo['name'].'"');

		readfile($fileInfo['path']);

		ob_end_flush();

		exit();
	}

	public function write($filename = NULL) {
		$timestamp = date('Ymd-hiA');
		$isRead = isset($this->filename);
		$extension = 'xls';

		if($isRead) {
			$extension = explode('.', $this->filename);
			$extension = strtolower(end($extension));
		}

		if(!isset($filename)) {
			$filename = 'tmp'.strtoupper(uniqid()).'-'.$timestamp.'.'.$extension;
		} else {
			if(preg_match('/\[date\]/i', $filename) > -1) {
				$date = date('Ymd');
				$filename = preg_replace('/\[date\]/i', $date, $filename);
			}

			if(preg_match('/\[time\]/i', $filename) > -1) {
				$time = date('hiA');
				$filename = preg_replace('/\[time\]/i', $time, $filename);
			}
		}

		$filename = $this->cache.$filename;

		$this->workbook->write(
			new java('java.io.FileOutputStream', $filename)
		);

		$this->filename = $filename;
		$this->saved = true;

		return $filename;
	}
}