<?php

namespace Boosty\POI;

class Cell extends POI {
	protected $workbook;

	public $index;
	public $type;
	public $style;
	public $value;

	private $cell;

	public function __construct($workbook, $intCell, $cell) {
		$this->workbook = $workbook;
		$this->cell = $cell;
		$this->index = $intCell;
		$this->type = self::TYPES[java_values($cell->getCellType())];
		$this->style = $this->getStyle();

		switch($this->type) {
			case 'STRING':
				$this->value = java_values($cell->getStringCellValue());
				break;
			case 'NUMERIC':
				$this->value = java_values($cell->getStringCellValue());
				break;
			case 'FORMULA':
				$this->value = java_values($cell->getCellFormula());
				break;
			case 'ERROR':
				$this->value = java_values($cell->getStringCellValue());
				break;
			case 'BOOLEAN':
				$this->value = java_values($cell->getBooleanCelLValue());
				break;
			case 'NULL':
				$this->setType('STRING');
				$this->value = '';
				break;
		}

		return $this;
	}

	public function getValue() {
		return $this->value;
	}

	public function getStyle() {
		$cellStyle = $this->cell->getCellStyle();
		$styles = array();

		return $styles;
	}

	public function setType($type) {
		$types = array_flip(self::TYPES);
		$intType = $types[$type];

		$this->cell->setCellType($intType);
		$this->type = $type;

		return $this;
	}

	public function setStyle($style = NULL) {
		if(is_string($style)) {
			$style = implode(';', $style);
		}

		if(is_array($style)) {
			$this->dump($style);
			exit;
		}

		$this->style = $style;

		return $this;
	}

	public function setValue($value) {
		if(is_numeric($value)) {
			$this->setType('NUMERIC');
			$value = doubleval($value);
		} else if(is_bool($value)) {
			$this->setType('BOOLEAN');
		} else {
			if(!strlen($value)) {

			}

			$this->setType('STRING');
		}

		$this->setSaved(false);
		$this->cell->setCellValue($value);
		$this->value = $value;

		return $this;
	}
}