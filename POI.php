<?php

namespace Boosty\POI;

use Exception as Exception,
	\kint AS kint;

class POI {
	const TYPES = array(
		'STRING',
		'NUMERIC',
		'FORMULA',
		'ERROR',
		'BOOLEAN',
		'NULL'
	);

	const _TYPES = array(
		'NUMERIC',
		'STRING',
		'FORMULA',
		'NULL',
		'BOOLEAN',
		'ERROR'
	);

	protected function dump($var) {
		kint::dump($var);
	}

	protected function exception($message) {
		echo '<pre>';
		throw new Exception($message);
		echo '</pre>';
	}

	protected function getFileInfo($file) {
		$mimeTypes = array(
			'xls' => 'application/vnd.ms-excel',
			'xlsx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
		);

		$name = explode('/', $file);
		$name = end($name);

		$extension = explode('.', $name);
		$extension = end($extension);

		$fileinfo = array(
			'name' => $name,
			'path' => $file,
			'extension' => $extension,
			'mimetype'=> $mimeTypes[$extension],
			'size' => filesize($file)
		);

		return $fileinfo;
	}

	protected function setSaved($isSaved) {
		if(isset($this->workbook)) {
			$reflection = new \ReflectionObject($this->workbook);

			$property = $reflection->getProperty('saved');
			$property->setAccessible(true);
			$property->setValue($this->workbook, $isSaved);
			$property->setAccessible(false);
		}
	}

	protected function checkMultiDimArray($array) {
		if(!is_array($array)) {
			echo '<pre>';
			$this->exception('The parameter needs to be an array!');
			echo '</pre>';
		} else {
			return (count($array) == count($array, COUNT_RECURSIVE)) ? false : true;
		}
	}
}