<?php

namespace Boosty\POI;

class Row extends POI {
	protected $workbook;

	public $index;

	private $row;
	private $cell;
	private $cells = array();

	public function __construct($workbook, $intRow, $row) {
		$this->workbook = $workbook;
		$this->index = $intRow;
		$this->row = $row;

		$cells = $row->cellIterator();

		// Read in physical CellsR
		/*while(java_values($cells->hasNext())) {
			$cell = $cells->next();
			$intCell = java_values($cell->getCellNum());
			$this->cells[$intRow] = new Cell($workbook, $intCell, $cell);
		}*/

		return $this;
	}

	public function createCell($intCell = NULL, $value = NULL) {
		$intCell = $this->checkCellIndex($intCell);
		$this->cells[$intCell] = new Cell($this->workbook, $intCell, $this->row->createCell($intCell));

		if(isset($value)) {
			$this->cells[$intCell]->setValue($value);
		}

		$this->setSaved(false);
		return $this->cells[$intCell];
	}

	public function getCell($intCell) {
		$intCell = $this->checkCellIndex($intCell);
		$this->cells[$intCell] = new Cell($workbook, $intCell, $this->row->getCell($intCell));
		return $this->cells[$intCell];
	}

	public function setCells($values) {
		if(is_array($values)) {
			$intCell = 0;

			foreach($values AS $value) {
				$intCell++;
				//$this->dump(array($intValue + 1, $value));
				$this->createCell($intCell, $value);
			}
		} else {
			echo '<pre>';
			$this->exception('The parameter needs to be an array!');
			echo '</pre>';
		}

		return $this;
	}

	private function checkCellIndex($intCell = NULL) {
		if(!isset($intCell)) {
			echo '<pre>';
			$this->exception('Please pass in the cell number!');
			echo '</pre>';
		} else {
			$intCell = $intCell - 1;
			if($intCell > -1) {
				return $intCell;
			} else {
				$this->exception('Cell numbers start at 1!');
			}
		}
	}
}