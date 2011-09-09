<?php
/*******************************************************************************
* PHPtoWord - Hello World example                                              *
*                                                                              *
* Version: 0.1                                                                 *
* Date:    2011-08-19                                                          *
* Author:  Miles Pickens                                                       *
*******************************************************************************/

	include 'phptoword.php';
	$ptw = new PhpToWord('Hello_Word');
	$ptw->open();
	$ptw->paragraph('Hello World');
	$ptw->paragraph('Hello World in a 2nd paragraph');
	$ptw->paragraph('Hello World in a 3rd paragraph');
	$ptw->pagebreak();
	//In Word the view the Print Preview to see this text on the 2nd page
	$ptw->paragraph('Hello World Printed on the 2nd page');
	$ptw->close();

