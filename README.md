ExcelComponent
==============

Excel component for CakePHP.  
Uses PHPExcel Library to proccess Microsoft Excel and CSV files.
    
    
    
Installing  
==========

1.  Install the latest version of PHPExcel in your app/Vendor folder in your CakePHP app, inside a PHPExcel folder. Tested with PHPExcel version 1.7.7, 2012-05-19  
The Vendor folder should end up like this:  
Vendor  
../ PHPExcel (folder)  
..../ PHPExcel (folder)  
..../ PHPExcel.php  
  
2.  Copy ExcelComponent.php class from this package's Controller/Component to your application's app/Controller/Component  
  
3.  Include ExcelComponent in your $components attribute in the controller your want to use.
```php
//Example:
public $components = array("Excel");
```