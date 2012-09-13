<?php
class ExcelException extends CakeException {
	
	public function __construct($message = null, $code = 500) {
		if (empty ( $message )) {
			$message = '';
		}
		parent::__construct ( $message, $code );
	}
}

?>