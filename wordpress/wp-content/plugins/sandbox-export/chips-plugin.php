<?php 
	/**
	* Plugin Name: WooCommerce To XLS
	* Description: Used to export WooCommerce order to XLS
	* Version: 0.1
	* Author: Group D
	*
	**/
	
	require_once plugin_dir_path( __FILE__ ) . 'Classes/PHPExcel.php';
	
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
					//For custom values meta key
					if($custom_val == true){
							$billing = array(
								'id' => $order->get_id(),
								'name' => $order->get_billing_first_name() . ' ' . $order->get_billing_last_name(),  
								'company' => $order->get_billing_company(),
								'address1' => $order->get_billing_address_1() . ' ' . $order -> get_billing_address_2(),
								'city' => $order->get_billing_city(),
								'post_code' => $order->get_billing_postcode(),
								'state' => $order->get_billing_state(),
								'email' => $order->get_billing_email(),
								'telephone'=> $order->get_billing_phone(),
								'product' => $item->get_name(),
								'qty' => $item->get_quantity(),
								'total' => $item->get_total(),
								'siswa1' => $item->get_meta('Nama Siswa (Anggota 1)'),
								'siswa2' => $item->get_meta('Nama Siswa (Anggota 2)')
							);
							array_push($data,$billing);
					}
					else{
							$billing = array(
								'name' => $order->get_billing_first_name() . ' ' . $order->get_billing_last_name(),  
								'company' => $order->get_billing_company(),
								'address1' => $order->get_billing_address_1() . ' ' . $order -> get_billing_address_2(),
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
			exportXLS($data);
		}
		
		function exportXLS($data){
			$objPHPExcel = new PHPExcel();	
			$sheet = $objPHPExcel->getActiveSheet();
			
			$sheet->setCellValue('A1', 'ID');
			$sheet->setCellValue('B1', 'Name');
			$sheet->setCellValue('c1', 'Company');
			$sheet->setCellValue('D1', 'Address');
			$sheet->setCellValue('E1', 'City');
			$sheet->setCellValue('F1', 'Postal Code');
			$sheet->setCellValue('G1', 'State');
			$sheet->setCellValue('H1', 'Email');
			$sheet->setCellValue('I1', 'Telephone');
			$sheet->setCellValue('J1', 'Product');
			$sheet->setCellValue('K1', 'Quantity');
			$sheet->setCellValue('L1', 'Total');
			$sheet->setCellValue('M1', 'Member 1');
			$sheet->setCellValue('N1', 'Member 2');
			$sheet->fromArray($data, NULL, 'A2');
			
			$cellIterator = $sheet->getRowIterator()->current()->getCellIterator();
		    $cellIterator->setIterateOnlyExistingCells(true);
		    foreach ($cellIterator as $cell) {
		        $sheet->getColumnDimension($cell->getColumn())->setAutoSize(true);
		    }
			
			$header = $objPHPExcel->getActiveSheet()->getStyle('A1:P1');
			$header->getFont()->setBold(true);
			$header->getFill()
			    ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
			    ->getStartColor()
			    ->setRGB('C0C0C0');
			
			header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
			header('Content-Disposition: attachment;filename="chips.xlsx"');
			header('Cache-Control: max-age=0');
			
			$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
			$objWriter->save('php://output');
			exit();
		}
	}
	
	
?>