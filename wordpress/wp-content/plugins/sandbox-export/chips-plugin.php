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
			};
			?>
			<div class="wrap">
				<h1><?= esc_html(get_admin_page_title()); ?></h1>
				<form action="<?php menu_page_url('full-orders-exporters') ?>" method="POST">
					<table>
						<tr>
							<th>
								Order Complete Date
							</th>
						</tr>
						<tr>
							<td>
								Date start
							</td>
							<td>
								:
							</td>
							<td>
								<input type="date" id="start" name="date-start" >
							</td>
						</tr>
						<tr>
							<td>
								Date end
							</td>
							<td>
								:
							</td>
							<td>
								<input type="date" id="end" name="date-end">
							</td>
						</tr>
						<tr>
							<td>
								Export All
							</td>
							<td>
								:
							</td>
							<td>
								<input type="checkbox" id="export-all" name="export-all">
							</td>
						</tr>
						<tr>
							<td>
								
							</td>
							<td>
								
							</td>
							<td style="text-align: right">
								<button type="submit" class="button">Export Now</button>
							</td>
						</tr>
					</table>
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
			$start = $_REQUEST['date-start'];
			$end = $_REQUEST['date-end'];
			$startUTC = strtotime($start . ' 00:00:00');
			$endUTC = strtotime($end . ' 23:59:59');
			if(!isset($_POST['export-all'])){
				if($start !== ""  && $end !== ""){
					$args = array(
						'order' => 'ASC',
						'limit' => 9999,
						'date_completed' => $startUTC . '...' . $endUTC,
					);
				}
				else{
					add_action('admin_notices','failed_notice');
					function failed_notice(){
						echo '<div class="notice notice-warning is-dismissible">
					             <p>Please pick a date.</p>
					         </div>';
					}
					return;
				}
			}
			else{
				$args = array(
					'order' => 'ASC',
					'limit' => 9999,
				);
			}
			
			$orders = WC_get_orders($args);
			if(!empty($orders)){
				$logger = wc_get_logger();
				$data = array();
				foreach($orders as $order){
					foreach($order->get_items() as $item_id => $item){
						$billing = array(
							'id' => $order->get_id(),
							'date' => $order->get_date_completed(),
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
				}
				//print_r($orders);
				$logger -> info(print_r($data,TRUE));
				exportXLS($data);
			}
			else{
				add_action('admin_notices','failed_notice');
					function failed_notice(){
						echo '<div class="notice notice-warning is-dismissible">
					             <p>No order exist.</p>
					         </div>';
					}
					return;
			}
			
		}
		
		function exportXLS($data){
			$objPHPExcel = new PHPExcel();	
			$sheet = $objPHPExcel->getActiveSheet();
			
			$sheet->setCellValue('A1', 'ID');
			$sheet->setCellValue('B1', 'Date Completed');
			$sheet->setCellValue('C1', 'Name');
			$sheet->setCellValue('D1', 'Company');
			$sheet->setCellValue('E1', 'Address');
			$sheet->setCellValue('F1', 'City');
			$sheet->setCellValue('G1', 'Postal Code');
			$sheet->setCellValue('H1', 'State');
			$sheet->setCellValue('I1', 'Email');
			$sheet->setCellValue('J1', 'Telephone');
			$sheet->setCellValue('K1', 'Product');
			$sheet->setCellValue('L1', 'Quantity');
			$sheet->setCellValue('M1', 'Total');
			$sheet->setCellValue('N1', 'Member 1');
			$sheet->setCellValue('O1', 'Member 2');
			$sheet->fromArray($data, NULL, 'A2');
			
			$cellIterator = $sheet->getRowIterator()->current()->getCellIterator();
		    $cellIterator->setIterateOnlyExistingCells(true);
		    foreach ($cellIterator as $cell) {
		        $sheet->getColumnDimension($cell->getColumn())->setAutoSize(true);
		    }
			
			
			$header = $objPHPExcel->getActiveSheet()->getStyle('A1:O1');
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