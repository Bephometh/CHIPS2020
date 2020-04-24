<?php 
	/**
	* Plugin Name: WooCommerce To XLS
	* Description: Used to export WooCommerce order to XLS
	* Version: 0.1
	* Author: Group D
	*
	**/

	if (!defined('ABSPATH')) {
		exit;
	}

	add_action('admin_menu', 'register_full_order_exporters_menu');

	function register_full_order_exporters_menu() {
		if(in_array('woocommerce/woocommerce.php', apply_filters('active_plugins', get_option('active_plugins')))) {
			add_submenu_page('woocommerce', 'Full Orders Exporter', 'Full Export Orders', 'manage_options', 'full-orders-exporters', '
			full_order_exporters_page') // TODO
		}

		function full_order_exporters_page() {
			//check user capabilities
			if(!current_user_can('manage_options')) { //TODO
				return;
			}
			?>
			<div class="wrap">
			<h1>TESTING JUDUL</h1>
			</div>
			<?php
		}
	}
