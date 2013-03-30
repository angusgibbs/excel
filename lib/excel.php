<?php

class Excel {
	private $doc;

	private $tableOptionsMap;
	private $rowOptionsMap;

	function __construct() {
		// Initialize the document to an empty string
		$this->doc = '';

		// Add the doctype
		$this->doc .= '<?xml version="1.0"?>' . "\n";

		// Add an opening <Workbook> tag
		$this->doc .= '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"' . "\n\t" .
			'xmlns:o="urn:schemas-microsoft-com:office:office"' . "\n\t" .
			'xmlns:x="urn:schemas-microsoft-com:office:excel"' . "\n\t" .
			'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' . "\n\t" .
			'xmlns:html="http://www.w3.org/TR/REC-html40">' . "\n";

		// Define the option maps
		$this->defineMaps();

		// Add the styles (so dates get displayed properly)
		$this->doc .= "\t" . '<Styles>' . "\n" .
			"\t\t" . '<Style ss:ID="s62">' . "\n" .
			"\t\t\t" . '<NumberFormat ss:Format="Short Date"/>' . "\n" .
			"\t\t" . '</Style>' . "\n" .
			"\t" . '</Styles>' . "\n";
	}

	public function sheet($name, $fn) {
		$this->doc .= "\t" . '<Worksheet ss:Name="' . $name . '">' . "\n";
		call_user_func($fn, $this);
		$this->doc .= "\t" . '</Worksheet>' . "\n";
	}

	public function table($fn, $options = array()) {
		// Open the open <Table> tag
		$this->doc .= "\t\t" . '<Table';

		// Add any attributes that are needed
		foreach ($options as $key => $value) {
			$name = $this->tableOptionsMap[$key];
			$this->doc .= ' ' . $name . '="' . $value . '"';
		}

		// Close the open <Table> tag
		$this->doc .= '>' . "\n";

		// Add in the table contents
		call_user_func($fn, $this);

		// Close the <Table> tag
		$this->doc .= "\t\t" . '</Table>' . "\n";
	}

	public function row($fn, $options = array()) {
		// Open the open <Row> tag
		$this->doc .= "\t\t\t" . '<Row';

		// Add any attributes that are needed
		foreach ($options as $key => $value) {
			$name = $this->rowOptionsMap[$key];
			$this->doc .= ' ' . $name . '="' . $value . '"';
		}

		// Close the open <Row> tag
		$this->doc .= '>' . "\n";

		// Add in the row contents
		call_user_func($fn, $this);

		// Close the <Row> tag
		$this->doc .= "\t\t\t" . '</Row>' . "\n";
	}

	public function cell($val, $type = null) {
		// Figure out the input type if it was not specified
		if (!$type) {
			$type = static::getType($val);
		}

		// Change the value to a proper timestamp if the type is DateTime
		if ($type == 'DateTime') {
			$val = date('Y-m-d\TH:i:s.u', strtotime($val));
		}

		// Add an open <Cell> tag
		$this->doc .= "\t\t\t\t" . '<Cell';

		// Add the DateTime style if this is a DateTime
		if ($type == 'DateTime') {
			$this->doc .= ' ss:StyleID="s62"';
		}

		// Finish the open <Cell> tag and open a <Data> tag
		$this->doc .= '><Data ss:Type="' . $type . '">';

		// Add the cell's value
		$this->doc .= $val;

		// Close the <Data> and <Cell> tags
		$this->doc .= '</Data></Cell>' . "\n";
	}

	public function fill($vals) {
		foreach ($vals as $val) {
			$this->cell($val);
		}
	}

	public function populate($table) {
		// Temporarily store the table data on the Excel object
		$this->tempTable = $table;

		// Construct the table
		$this->table(function($table) {
			// Add each row
			foreach ($table->tempTable as $vals) {
				// Temporarily store the values on the Excel object
				$table->tempVals = $vals;

				// Fill the row with the contents
				$table->row(function($row) {
					$row->fill($row->tempVals);
				});

				// Remove the row data from the Excel object
				unset($table->tempVals);
			}
		});

		// Remove the table data from the Excel object
		unset($this->tempTable);
	}

	private static function getType($var) {
		if (is_bool($var)) return 'Boolean';
		if (is_numeric($var)) return 'Number';
		if (strtotime($var)) return 'DateTime';
		return 'String';
	}

	public function __toString() {
		return $this->doc;
	}

	public function end() {
		// Add the closing <Workbook> tag
		$this->doc .= '</Workbook>' . "\n";

		return $this->doc;
	}

	private function defineMaps() {
		$this->tableOptionsMap = array(
			'colWidth' => 'DefaultColumnWidth',
			'rowHeight' => 'DefaultRowHeight',
			'colCount' => 'ExpandedColumnCount',
			'rowCount' => 'ExpandedRowCount',
			'left' => 'LeftCell',
			'top' => 'TopCell'
		);

		$this->rowOptionsMap = array(
			'hidden' => 'Hidden',
			'height' => 'Height'
		);
	}
}
