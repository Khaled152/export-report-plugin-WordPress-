<?php
/*
Plugin Name: Export Courses to Excel
Description: Plugin to export courses data to Excel using PHPSpreadsheet
Version: 1.0
Author: Your Name
*/

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require_once plugin_dir_path(__FILE__) . 'vendor/autoload.php';

// Function to execute SQL query and export results to Excel
function export_courses_to_excel() {
    global $wpdb;

    // Check user capabilities
    if (!current_user_can('manage_options')) {
        wp_die('You do not have sufficient permissions to access this page.');
    }

    // Check if the query is set
    if (!isset($_POST['custom_query'])) {
        wp_die('No query provided.');
    }

    // Get the custom SQL query from the form
    $query = stripslashes(sanitize_text_field($_POST['custom_query']));

    // Ensure that the query is a SELECT statement
    if (stripos(trim($query), 'SELECT') !== 0) {
        wp_die('Only SELECT queries are allowed.');
    }

    // Additional security check to ensure the query does not contain dangerous commands
    $disallowed = ['INSERT', 'UPDATE', 'DELETE', 'DROP', 'ALTER'];
    foreach ($disallowed as $cmd) {
        if (stripos($query, $cmd) !== false) {
            wp_die('Disallowed SQL command found in the query.');
        }
    }

    // Get results from the database
    $results = $wpdb->get_results($query);

    if (empty($results)) {
        wp_die('No results found or query error.');
    }

    // Create a new Spreadsheet object
    $spreadsheet = new Spreadsheet();

    // Get the active sheet
    $sheet = $spreadsheet->getActiveSheet();

    // Set headers
    $headers = array_keys((array)$results[0]);
    $col = 1;
    foreach ($headers as $header) {
        $sheet->setCellValueByColumnAndRow($col, 1, $header);
        $col++;
    }

    // Add data to the Excel file
    $row = 2;
    foreach ($results as $result) {
        $col = 1;
        foreach ($result as $value) {
            $sheet->setCellValueByColumnAndRow($col, $row, $value);
            $col++;
        }
        $row++;
    }

    // Set headers for download
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="courses.xlsx"');
    header('Cache-Control: max-age=0');

    // Output the Excel file to browser
    $writer = new Xlsx($spreadsheet);
    $writer->save('php://output');

    exit;
}

// Function to render export page content
function render_export_page() {
  ?>
    <div class="wrap">
        <h1>Export Courses to Excel</h1>
        <form method="post">
            <?php wp_nonce_field('export_courses_nonce', 'export_courses_nonce'); ?>
            <input type="hidden" name="export_courses" value="1" />
            <textarea name="custom_query" cols="50" rows="10" placeholder="Enter your custom SQL query"></textarea>
            <?php submit_button('Export Courses'); ?>
        </form>
    </div>
  <?php
}

// Hook to handle export process when form is submitted
function handle_export_process() {
    if (isset($_POST['export_courses']) && isset($_POST['export_courses_nonce']) && wp_verify_nonce($_POST['export_courses_nonce'], 'export_courses_nonce')) {
        export_courses_to_excel();
    }
}

// Hook to add admin page
function add_export_page() {
    add_menu_page(
        'Export report',
        'Export report',
        'manage_options',
        'export-report',
        'render_export_page'
    );
}

add_action('admin_menu', 'add_export_page');
// Hook to handle export process
add_action('admin_init', 'handle_export_process');
?>
