# Excel by [Angus Gibbs](http://angusgibbs.com)

## License

Excel is licensed under the MIT license.

## Usage

Simply require `excel.php` ([raw](https://raw.github.com/angusgibbs/excel/master/lib/excel.php)) and you'll be provided with a global `Excel` object that you can work with.

## API

### Overall example

```php
$book = new Excel;
$book->sheet('Sheet1', function($sheet) {
	$sheet->table(function($table) {
		$table->row(function($row) {
			$row->cell('Name');
			$row->cell('Age');
			$row->cell('Birth Date');
		});

		$table->row(function($row) {
			$row->fill(array('John', 24, '4/5/2006'));			
		});

		$table->row(function($row) {
			$row->fill(array('Joe', 20, '5/6/2007'));
		});
	});
});
```

### Creating a new sheet

Call `sheet` on your `Excel` object. `sheet` takes the name of the sheet and a closure that will define the contents.

```php
$book->sheet(function($sheet) {
	// Do stuff with $sheet
});
```

### Creating a new table

Call `table` on the first parameter supplied by the `sheet` closure. `table` takes a closure that will define the table contents.

```php
$book->sheet(function($sheet) {
	$sheet->table(function($table) {
		// Do stuff with $table
	});
});
```

### Creating a row

Call `row` on the first parameter supplied by the `table` closure. `row` takes a closure that will define the row contents.

```php
$book->sheet(function($sheet) {
	$sheet->table(function($table) {
		$table->row(function($row) {
			// Do stuff with $row
		});
	});
});
```

### Adding a data value (cell)

Call `cell` on the first parameter supplied by the `row` closure. `cell` takes the value as the first parameter and an optional second parameter, the data type. The data type must be one of 'String', 'Number', or 'DateTime'.

```php
$row->cell('Angus');
$row->cell(16); // Data type of Number will be guessed
$row->cell('16'); // Data type of Number will still be guessed
$row->cell(16, 'String'); // Override the default data type
```

### Filling a row with values

If you simply wish to fill a row with values from an array, you can call `$row->fill(arr)` inside the row closure to add each element on the array. *Note: data types will be automatically set.*

```php
$table->row(function($row) {
	$row->fill(array('Angus', 'Gibbs', 'Hello'));
});
```

### Populating a sheet with data

If you simply wish to populate a sheet with data from a two dimensional array, you can do that with `$sheet->populate`.

```php
$book = new Excel;
$book->sheet('Sheet1', function($sheet) {
	$sheet->populate(array(
		array('Name', 'Age'),
		array('Person', 30),
		array('Joe', 25),
		array('Bob', 65)
	));
});
```
