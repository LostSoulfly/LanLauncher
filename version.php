<?php 
$logfile= 'GetLog.txt'; 
$IP = $_SERVER['REMOTE_ADDR']; 
$logdetails=  date("F j, Y, g:i a").' with Version: '. $_GET['ver'] .' From: ('. $_SERVER['REMOTE_ADDR'].') Rnd: ('.$_GET['rnd'].')'; 
$fp = fopen($logfile, "a");  
fwrite($fp, $logdetails); 
fwrite($fp, "\n"); 
fclose($fp); 
Readfile('Version.txt');
#Readfile('http://dl.dropbox.com/u/4275989/Version.txt');
?>