<?php
// Before removing this file, please verify the PHP ini setting `auto_prepend_file` does not point to this.

if (file_exists('/home/chips/wordpress/wp-content/plugins/wordfence/waf/bootstrap.php')) {
	define("WFWAF_LOG_PATH", '/home/chips/wordpress/wp-content/wflogs/');
	include_once '/home/chips/wordpress/wp-content/plugins/wordfence/waf/bootstrap.php';
}
?>