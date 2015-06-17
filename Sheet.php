<?php

namespace Boosty\POI;

class Sheet extends POI {
	public $name;

	protected $workbook;
	private $sheet;
	private $rows;

	public function __construct($workbook, $name, $sheet) {
		$this->workbook = $workbook;
		$this->name = $name;
		$this->sheet = $sheet;

		$rows = $sheet->rowIterator();

		return $this;
	}

	public function createRow($intRow = NULL) {
		$intRow = $this->checkRowIndex($intRow);
		$this->rows[$intRow] = new Row($this->workbook, $intRow, $this->sheet->createRow($intRow));
		$this->setSaved(false);
		return $this->rows[$intRow];
	}

	public function getRow($intRow) {
		$intRow = $this->checkRowIndex($intRow);
		$this->rows[$intRow] = new Row($workbook, $intRow, $this->sheet->getRow($intRow));
		return $this->rows[$intRow];
	}

	public function getTableData() {
		$XSSFTables = $this->sheet->getTables()->iterator();
		$dateUtil = java('org.apache.poi.ss.usermodel.DateUtil');
		$tables = array();

		// Fetch Excel Sheet Tables
		while(java_values($XSSFTables->hasNext())) {
			$XSSFTable = $XSSFTables->next();
			$name = java_values($XSSFTable->getName());

			$startCellReference = $XSSFTable->getStartCellReference();
			$endCellReference = $XSSFTable->getEndCellReference();

			$tables[$name] = array(
				'rows' => array(
					java_values($startCellReference->getRow()),
					java_values($endCellReference->getRow())
				),

				'columns' => array(
					java_values($startCellReference->getCol()),
					java_values($endCellReference->getCol())
				)
			);
		}

		// Fetch all physical cell data
		$dataset = array();
		$rows = $this->sheet->rowIterator();
		$intPrevRow = 0;

		while(java_values($rows->hasNext())) {
			$row = $rows->next();
			$intRow = java_values($row->getRowNum());

			if($intRow - $intPrevRow > 1) {
				$intDeltaRow = ($intRow - $intPrevRow - 1);

				while($intDeltaRow--) {
					$dataset[] = array();
				}
			}

			$cells = $row->cellIterator();

			$dataset[$intRow] = array();
			$intPrevCol = 0;

			while(java_values($cells->hasNext())) {
				$cell = $cells->next();
				$intCol = java_values($cell->getColumnIndex());
				$cellType = java_values($cell->getCellType());
				$isNumeric = ($cellType == 0);
				$isDate = false;

				// Parse if formula result is numeric type
				if($cellType == 2) {
					$cellFormatType = java_values($cell->getCachedFormulaResultType());

					switch($cellFormatType) {
						// Numeric Type
						case 0:
							$cell->setCellType(0);
							$isNumeric = true;
							break;
						// Errors
						case 5:
							$cell->setCellType(3);
							$isNumeric = false;
							break;
					}
				}

				// For Cell errors, set to blank value
				if($cellType == 5) {
					$cell->setCellType(3);
					$isNumeric = false;
				}

				if($isNumeric) {
					$cell->setCellType(0);
					$isDate = java_values($dateUtil->isCellInternalDateFormatted($cell)) OR java_values($dateUtil->isCellDateFormatted($cell));
				} else {
					$cell->setCellType($cellType);
				}

				if($intCol - $intPrevCol > 1) {
					$intDeltaCol = ($intCol - $intPrevCol - 1);

					while($intDeltaCol--) {
						$dataset[$intRow][] = NULL;
					}
				}

				if($isDate) {
					$cellValue = java_values($cell->getDateCellValue()->toString());
				} else {
					$cell->setCellType(1);
					$cellValue = strval(java_values($cell->getStringCellValue()));
					$cellValue = strlen($cellValue) ? $cellValue : NULL;

					if($isNumeric && $cellValue != NULL) {
						$cellValue = doubleval($cellValue);
					}
				}

				$dataset[$intRow][] = $cellValue;
				$intPrevCol = $intCol;
			}

			$intPrevRow = $intRow;
		}

		// Slice and Dice the resultset
		$data = array();

		foreach($tables AS $tableName=>$table) {
			$results = $this->slice(
				$dataset,
				$table['rows'][0],
				($table['rows'][1] - $table['rows'][0]) + 1,
				$table['columns'][0],
				($table['columns'][1] - $table['columns'][0]) + 1
			);

			$columns = array_shift($results);

			foreach($columns AS $intCol=>$column) {
				$columns[$column] = $intCol;
				unset($columns[$intCol]);
			}

			$data[$tableName] = array(
				'columns' => $columns,
				'data' => $results
			);
		}

		return $data;
	}

	public function setColumnAutoWidth($intCol) {
		$this->checkRowIndex($intCol);
		$this->sheet->autoSizeColumn($intCol);

		return $this;
	}

	private function slice($array, $rowOffset, $rowLength, $columnOffset, $columnLength) {
		$resultset = array_slice($array, $rowOffset, $rowLength);

		foreach($resultset AS &$row) {
			$row = array_slice($row, $columnOffset, $columnLength);
		}

		return $resultset;
	}

	private function checkRowIndex($intRow = NULL) {
		if(!isset($intRow)) {
			$this->exception('Please pass in the row number!');
		} else {
			$intRow = $intRow - 1;
			if($intRow > -1) {
				return $intRow;
			} else {
				$this->exception('Row numbers start at 1!');
			}
		}
	}

// 	public function createFreezePane();
}