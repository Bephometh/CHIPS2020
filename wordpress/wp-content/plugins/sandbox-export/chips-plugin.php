<?php 
	/**
	* Plugin Name: WooCommerce To XLS
	* Description: Used to export WooCommerce order to XLS
	* Version: 0.1
	* Author: Group D
	*
	**/

	if ( ! defined( 'ABSPATH' ) ) {
		exit; // Exit if accessed directly
	}

	if( ! in_array( 'woocommerce/woocommerce.php', apply_filters( 'active_plugins', get_option( 'active_plugins' )))) {
		exit;
	}

	add_action('admin_menu', 'register_full_order_exporters_menu');

	function register_full_order_exporters_menu() {
		$hook_name = add_submenu_page('woocommerce', 'Full Orders Exporter', 'Full Export Orders', 'view_woocommerce_reports', 'full-orders-exporters', 'full_order_exporters_page');
		function full_order_exporters_page() {
			//check user capabilities
			if( ! current_user_can('view_woocommerce_reports')) {
				return;
			}
			?>
			<div class="wrap">
				<h1><?= esc_html(get_admin_page_title()); ?></h1>
				<form action="<?php menu_page_url('full-orders-exporters') ?>" method="POST">
					<button type="submit" class="button">Export Now</button>
				</form>
                <p>*Export all orders(CHIPS participant) to Excel</p>
			</div>
			<?php
		}
		add_action('load-' .  $hook_name, 'full_order_exporters_page_submit');

		function full_order_exporters_page_submit() {
			$logger = wc_get_logger();
			$logger -> info('Export request receiver');

		}
	}
?>