ExcelComponent
==============

Excel component for CakePHP.  
Uses PHPExcel Library to proccess Microsoft Excel and CSV files.
    
    
    
Installing  
==========

-Install the latest version of PHPExcel in your app/Vendor folder in your CakePHP app, inside a PHPExcel folder.  
The Vendor folder should end up like this:  
Vendor  
..| PHPExcel (folder)  
....| PHPExcel (folder)  
....| PHPExcel.php  
  
-Copy ExcelComponent.php class from this package's Controller/Component to your application's app/Controller/Component  
  
-Include ExcelComponent in your $components attribute in the controller your want to use. Example: public $components = array("Excel");