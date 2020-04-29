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
		// add_action('load-' .  $hook_name, 'full_order_exporters_page_submit');
		if('POST' === $_SERVER['REQUEST_METHOD']) {
			add_action('load-' .  $hook_name, 'getOrders');
		}
		
		function full_order_exporters_page_submit() {
			if('POST' === $_SERVER['REQUEST_METHOD']) {
				$logger = wc_get_logger();
				$logger -> info('Export request receiver');
			}
		}
		
		function getOrders(){
			$custom_val = true;
			$args = array(
				'order' => 'ASC',
				'limit' => 9999
			);
			$query = new WC_Order_query($args);
			$orders = WC_get_orders($args);
			$logger = wc_get_logger();
			$data = array();
			foreach($orders as $order){
				foreach($order->get_items() as $item_id => $item){
					//For custome values meta key
					if($custom_val == true){
							$billing = array(
								'id' => $order->get_id(),
								'name' => $order->get_billing_first_name(),
								'last_name' =>$order->get_billing_last_name(),
								'company' => $order->get_billing_company(),
								'address1' => $order->get_billing_address_1(),
								'address2' =>$order -> get_billing_address_2(),
								'city' => $order->get_billing_city(),
								'post_code' => $order->get_billing_postcode(),
								'state' => $order->get_billing_state(),
								'email' => $order->get_billing_email(),
								'telephone'=> $order->get_billing_phone(),
								'product' => $item->get_name(),
								'qty' => $item->get_quantity(),
								'total' => $item->get_total(),
								'metas' => ''
							);
							$metas = array(
								'Nama Siswa (Anggota 1)' =>  $item->get_meta('Nama Siswa (Anggota 1)'),
								'Nama Siswa (Anggota 2)' => $item->get_meta('Nama Siswa (Anggota 2)')
							);
							$billing['metas'] = $metas;
							array_push($data,$billing);
					}
					else{
							$billing = array(
								'name' => $order->get_billing_first_name(),
								'last_name' =>$order->get_billing_last_name(),
								'company' => $order->get_billing_company(),
								'address1' => $order->get_billing_address_1(),
								'address2' =>$order -> get_billing_address_2(),
								'city' => $order->get_billing_city(),
								'post_code' => $order->get_billing_postcode(),
								'state' => $order->get_billing_state(),
								'email' => $order->get_billing_email(),
								'telephone'=> $order->get_billing_phone(),
								'product' => $item->get_name(),
								'qty' => $item->get_quantity(),
								'total' => $item->get_total()
							);
							 array_push($data,$billing);
					}
				}
			}
			$logger -> info(print_r($data,TRUE));
			//return $data;	
			exportXLS($data);
		}
		
		function exportXLS($data){
			header("Content-type: application/vnd-ms-excel");
			header("Content-Disposition: attachment; filename= chips.xls");
			echo "<table>";
			foreach($data as $user){
				echo "<tr>";
				foreach($user as $value){
					if(!is_array($value)){
						echo "<td>";
						echo is_array($value);
						echo print_r($value,TRUE);
						echo "</td>";
					}
					else{
						foreach($value as $arr_value){
							echo "<td>";
							echo print_r($arr_value,TRUE);
							echo "</td>";
						}
					}
				}
				echo "</tr>";
			}
			echo "</table>";
			exit();
		}
	}
	
	
?>