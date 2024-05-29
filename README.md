# Export Reports to Excel

## Description

The "Export Reports to Excel" plugin enables administrators to easily export detailed reports from their WordPress site to an Excel file. Utilizing the PHPSpreadsheet library, this plugin executes a comprehensive SQL query to gather report data and generates a downloadable Excel file. Perfect for administrators who need to manage and analyze report data offline.

## Key Features

- **Effortless Export:** Export comprehensive reports, including user details, enrollment information, completion statuses, and assessment results, directly to an Excel file.
- **Customizable SQL Query:** Tailor the SQL query to match your specific reporting needs.
- **Access Control:** Ensures that only users with the appropriate permissions (e.g., administrators) can export the data.
- **User-Friendly Interface:** Provides an intuitive admin page in the WordPress dashboard to initiate the export process.
- **Instant File Download:** Generates the Excel file on-the-fly and prompts the user to download it immediately.

## Version

1.0

## Author

Khaled Ahmed 

## Installation

1. Download and install Composer if you haven't already.
2. Run `composer require phpoffice/phpspreadsheet` in the root directory of your plugin to install PHPSpreadsheet.
3. Upload the plugin files to the `/wp-content/plugins/export-reports-to-excel` directory, or install the plugin through the WordPress plugins screen directly.
4. Activate the plugin through the 'Plugins' screen in WordPress.
5. Navigate to the "Export Reports" page in the WordPress admin menu.

## Usage

1. **Navigate to the Export Page:** Go to the "Export Reports" page in the WordPress admin menu.
2. **Export Data:** Click the "Export Reports" button to generate and download the Excel file with your report data.

## Customizing the SQL Query

To use this plugin for your specific reporting needs, you must customize the SQL query in the `export_courses_to_excel` function. The default query is a placeholder and needs to be replaced with your own query to gather the required report data.

1. Open the `export-reports-to-excel.php` file.
2. Locate the following line:

    ```php
    $query = 'YOUR_SQL_QUERY_HERE';
    ```

3. Replace `'YOUR_SQL_QUERY_HERE'` with your actual SQL query. Ensure your query retrieves the necessary fields and data for your report.

Example:

```php
$query = 'SELECT user_login AS `User Login`, user_email AS `User Email`, display_name AS `Display Name` FROM wp_users';
