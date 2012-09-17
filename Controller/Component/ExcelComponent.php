<?php

/**
 * ExcelComponent is a PHPExcel wrapper for CakePHP with some helper methods.
 * For instance, converting an excel worksheet (or part of a worksheet) to a CakePHP data array ready to be used inside Model::save()
 * 
 * @version 0.0.1 - First Release
 * @author Caio F. Landau
 * 
 */

App::uses("ExcelException", "Lib/Error");

class ExcelComponent extends Component {
	const ERR_NO_FILE_LOADED = 'No file or invalid file loaded';
	const ERR_INVALID_WORKSHEETS_ARRAY = 'Invalid worksheets array';
	
	private $workingFile = null;
	
	/**
	 * When modifying the ExcelComponent class, you should not access this directly. Use getReader() method instead.
	 * @var PHPExcel
	 */
	private $reader = null;
	
	/**
	 * (non-PHPdoc)
	 * @see Component::initialize()
	 */
	public function initialize(Controller $controller) {
		App::import('Vendor', 'PHPExcel/PHPExcel');
	}
	
	/**
	 * Sets the file for the component to work with
	 * @param string $filePath the path of the Excel/CSV file to load
	 * 
	 * @return void
	 */
	public function setWorkingFile($filePath, $absolute = false) {
		$this->unsetWorkingFile();
		if (!$absolute) {
			$filePath = WWW_ROOT.$filePath;
		}
		$this->workingFile = $filePath;
	}
	
	/**
	 * Unsets the working file. Releases memory used by the PHPExcel reader
	 * 
	 * @return void
	 */
	public function unsetWorkingFile() {
		$this->workingFile = null;
		$this->reader = null;
	}
	
	/**
	 * Returns a PHPExcel object of $filePath provided
	 * @param string $filePath
	 * @return PHPExcel the PHPExcel object containing the file
	 */
	public function getObjectForFile($filePath) {
		$inputFileType = PHPExcel_IOFactory::identify($filePath);
		$objReader = PHPExcel_IOFactory::createReader($inputFileType);
		$objPHPExcel = $objReader->load($filePath);
		return $objPHPExcel;
	}
	
	/**
	 * Gets currently loaded file as an array
	 * @return array The array containing file data
	 * @return boolean false if no file is loaded
	 */
	public function getAllSheetsAsArray() {
		if ($this->getReader() != null) {
			$sheets = $this->getReader()->getAllSheets();
			$return = array();
			foreach($sheets as $sheet) {
				$return[]['Sheet'] = $sheet->toArray(null, true, false, true);
			}
			return $return;
		}
		else {
			throw new ExcelException(ExcelComponent::ERR_NO_FILE_LOADED);
			return false;
		}
	}
	
	/**
	 * Finds data in the file. Looks inside cells. Similar to CakePHP's Model::find() method, but returns an array of cell coordinates in which the value was found
	 * 
	 * Find conditions:
	 * <pre>
	 * array(
	 * 	"text" => array("some search") //What to look for (string or an array of strings). If you provide an array of strings, it will search for ANY ocurrence. Regex syntax supported.
	 * 	"case_sensitive" => false //Should the search be case sensitive? Defaults to false (case-insensitive). If providing a regex for "text", this should always be "false".
	 * )
	 * </pre>
	 * 
	 * @param string $type Supports 'all', 'first' and 'count'
	 * @param string|array $conditions The find conditions. See method description for details.
	 * @param array $worksheets An array defining which worksheets to look in. Can be a numeric array (worksheet numbers) or string array (worksheet names). Defaults to all worksheets
	 * 
	 * @return array Coordinates on which the value was found
	 */
	public function find($type, $conditions, $worksheets = null) {
		$active_sheet = $this->getReader()->getIndex($this->getReader()->getActiveSheet());
		if ($this->getReader() == null) {
			throw new ExcelException(ExcelComponent::ERR_NO_FILE_LOADED);
			return false;
		}
		if ($worksheets != null) {
			if (!is_array($worksheets)) {
				$worksheets = array($worksheets);
			}
			if (!is_numeric($worksheets[0])) {
				$found = $this->findBySheetsName($type, $conditions, $worksheets);
			}
			else {
				$found = $this->findBySheetsNumber($type, $conditions, $worksheets);
			}
		}
		else {
			$found = $this->findInAllSheets($type, $conditions);
		}
		//Set active sheet back to what it was before find() was called
		$this->getReader()->setActiveSheetIndex($active_sheet);
		return $found;
	}
	
	/**
	 * Finds data in the active worksheet. Looks inside cells. Similar to CakePHP's Model::find() method, but returns an array of cell coordinates in which the value was found
	 * 
	 * Find conditions:
	 * <pre>
	 * array(
	 * 	"text" => array("some search"), //What to look for (string or an array of strings). If you provide an array of strings, it will search for ANY ocurrence. Regex syntax supported.
	 * 	"case_sensitive" => false //Should the search be case sensitive? If providing a regex for "text", this should always be "false". Defaults to false (case-insensitive).
	 * )
	 * </pre>
	 * 
	 * @param string $type Supports 'all', 'first' and 'count' 
	 * @param string|array $conditions The find conditions. See method description for details.
	 * 
	 * @return array Array of cell coordinates and text found
	 */
	public function findInActiveSheet($type, $conditions) {
		$raw_data = $this->getReader()->getActiveSheet()->toArray(null, true, false, true);
		$found = array();
		if (!is_array($conditions['text']))
			$conditions['text'] = array($conditions['text']);
		if (isset($conditions['case_sensitive']) && $conditions['case_sensitive']) {
			$cs = true;
		}
		else {
			$cs = false;
		}
		foreach($raw_data as $rowNum => $row) {
			foreach($row as $colCod => $cell) {
				foreach($conditions['text'] as $condition) {
					if ($cs) {
						$text = $condition;
						$cellVal = $cell;
					}
					else {
						$text = strtolower($condition);
						$cellVal = strtolower($cell);
					}
					if (preg_match('/'.$text.'/', $cellVal)) {
						$found[] = array("row" => $rowNum, 'col' => $colCod, 'cell' => $cell);
						if ($type == 'first')
							return $found;
						break;
					}
				}
			}
		}
		if (empty($found)) {
			if ($type == 'count')
				$found = 0;
			else
				$found = false;
		}
		else {
			if ($type == 'count') {
				$found = count($found);
			}
		}
		return $found;
	}
	
	/**
	 * Converts a worksheet's rows and columns to a data array ready to be used for Model::save(), Model::create() or Model::set().
	 *
	 * The format array should be in the format:
	 * 	array(
	 * 	    "Row or Column" => "Model.field" //See examples below
	 * 	)
	 *
	 * For example, to extract each row as an entry and each column as a field:
	 * <pre>
	 * 	array(
	 * 		"A" => "User.username", //Column A contains an User.username on each row
	 * 		"B" => "User.password", //Column B contains an User.password on each row
	 * 		"C" => "User.name" //Column C contains an User.name on each row
	 * 	)
	 * 
	 * //In this example, each row in the excel file represents a different User.
	 * </pre>
	 *
	 * -
	 *
	 * To use columns as entries and rows as fields:
	 * <pre>
	 * 	array(
	 * 		1 => "User.username", //Row 1 contains an User.username on each column
	 * 		2 => "User.password", //Row 2 contains an User.password on each column
	 * 		3 => "User.name" //Row 3 contains an User.name on each column
	 * 	)
	 * //In this example, each column in the excel file represents a different User.
	 * </pre>
	 *
	 * @param array $format Format array defining how to convert rows and columns from the worksheet
	 * @param string $range The range to get data from the worksheet. For example: 'A1:F10'. Defaults to entire worksheet
	 * @param string/int $worksheet Which worksheet to use (name or number). Defaults to currently active worksheet
	 *
	 * @return array The data array ready to be passed to Model::save()
	 */
	public function toDataArray($format = array(), $range = null, $worksheet = null) {
		if (!is_null($worksheet)) {
			if (!is_numeric($worksheet)) {
				$this->getReader()->setActiveSheetIndexByName($worksheet);
			}
			else {
				$this->getReader()->setActiveSheetIndex($worksheet);
			}
		}
		if ($range == null)
			$raw_array = $this->getReader()->getActiveSheet()->toArray(null, true, false, true);
		else
			$raw_array = $this->getReader()->getActiveSheet()->rangeToArray($range, null, true, false, true);
	
		$data_array = $this->rawArrayToDataArray($raw_array, $format);
		return $data_array;
	}
	
	/**
	 * Lazy loading the PHPExcel reader property.
	 * Access this from outside ExcelComponent to get access to the raw PHPExcel object for the currently loaded file.
	 * 
	 * @return PHPExcel reader
	 */
	public function getReader() {
		if(is_null($this->workingFile) || empty($this->workingFile)) {
			throw new ExcelException(ExcelComponent::ERR_NO_FILE_LOADED);
			return null;
		}
	
		if (is_null($this->reader)) {
			$this->reader = $this->getObjectForFile($this->workingFile);
		}

		return $this->reader;
	}
	
	/**
	 * @access private 
	 */
	private function findBySheetsName($type, $conditions, $worksheets) {
		foreach($worksheets as $worksheet) {
			$this->getReader()->setActiveSheetIndexByName($worksheet);
			$found[] = $this->findInActiveSheet($type, $conditions);
			if ($type == 'first' && !empty($found))
				return $found[0];
		}
		return $found;
	}
	/**
	 * @access private
	 */
	private function findBySheetsNumber($type, $conditions, $worksheets) {
		foreach($worksheets as $worksheet) {
			$this->getReader()->setActiveSheetIndex($worksheet);
			$found[$worksheet] = $this->findInActiveSheet($type, $conditions);
			if ($type == 'first' && !empty($found))
				return $found[0];
		}
		return $found;
	}
	/**
	 * @access private
	 */
	private function findInAllSheets($type, $conditions) {
		$allSheets = $this->getReader()->getAllSheets();
		foreach($allSheets as $sheet) {
			$this->getReader()->setActiveSheetIndex($this->getReader()->getIndex($sheet));
			$found[$this->getReader()->getIndex($sheet)] = $this->findInActiveSheet($type, $conditions);
		}
		return $found;
	}
	
	
	/**
	 * Converts a raw array (from PHPExcel_Worksheet::toArray()) to a CakePHP data array.
	 * 
	 * @param array $raw_array The raw array
	 * @param array $format The format to convert
	 * 
	 * @access private
	 */
	private function rawArrayToDataArray($raw_array, $format) {
		reset($format);
		if (!is_numeric(key($format))) {
			//Rows to entries
			$array = $this->rowsToEntries($raw_array, $format);
		}
		else {
			//Columns to entries
			$array = $this->columnsToEntries($raw_array, $format);
		}
		
		$array = $this->dotNotationToDataArray($array);
		return $array;
	}
	
	/**
	 * Excel rows to dot notation data array entries
	 * 
	 * @param array $raw_array
	 * @param array $format
	 * 
	 * @access private
	 */
	private function rowsToEntries($raw_array, $format) {
		$cont = 0;
		foreach($raw_array as $k => $row) {
			foreach($row as $col => $cell) {
				if (isset($format[$col])) {
					$array[$cont][$format[$col]] = $cell;
				}
			}
			$cont++;
		}
		return $array;
	}
	
	/**
	 * Excel columns to dot notation data array entries
	 *
	 * @param array $raw_array
	 * @param array $format
	 * 
	 * @access private
	 */
	private function columnsToEntries($raw_array, $format) {
		$x = 0;
		foreach ($raw_array as $k => $row) {
			foreach($row as $col => $valor) {
				if (isset($format[$k])) {
					$data[$format[$k]][] = $valor;
				}
			}
			$x++;
		}
		
		foreach($data as $campo => $valores) {
			foreach($valores as $index => $valor) {
				$retorno[$index][$campo] = $valor;
			}
		}
		return $retorno;
	}
	
	/**
	 * @access private 
	 */
	private function dotNotationToDataArray($array) {
		foreach($array as $x => $row) {
			foreach($row as $field => $value) {
				$tmp = explode(".", $field);
				$model = $tmp[0];
				$field = $tmp[1];
				$retorno[$x][$model][$field] = $value;
			}
		}
		return $retorno;
	}
	
	/**
	 * @access private
	 */
	private function validateWorksheetsArray($array) {
		foreach($array as $worksheet) {
			if (!is_numeric($worksheet)) {
				throw new ExcelException(ExcelComponent::ERR_INVALID_WORKSHEETS_ARRAY);
				return false;
			}
		}
	}
	
}

?>