<?php
/**This code fetches the product list from the JSON file,
 *  processes the data, and generates an HTML table and an Excel file for each product. */

 
// Include the necessary libraries
require_once '../vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Specify the relative path to the JSON file
$jsonFilePath = 'response.json';

// Read the contents of the JSON file
$response = file_get_contents($jsonFilePath);

// Decode the JSON data
$jsonData = json_decode($response, true);

// Check if decoding was successful
if ($jsonData === null || !isset($jsonData['data']) || !is_array($jsonData['data'])) {
    echo 'Error decoding JSON response or invalid data structure.';
} else {
    // Create an HTML table header
    $htmlTable = "<table border='1'>
        <tr>
            <th>Product SKU</th>
            <th>Product Name</th>
            <th>Product Price Before Discount</th>
            <th>Product Price After Discount</th>
            <th>Quantity</th>
            <th>Categories Names</th>
            <th>Categories Ids</th>
            <th>Promotion Title</th>
            <th>Metadata Title</th>
            <th>Metadata Description</th>
            <th>Description</th>
            <th>Product Image URLs</th>
        </tr>";

    // Iterate over each product
    foreach ($jsonData['data'] as $productInfo) {
        // Extract and process product information
        $productId = $productInfo['id'];
        $productName = $productInfo['name'];
        $productPriceBeforeDiscount = $productInfo['regular_price']['amount'];
        $productPriceAfterDiscount = $productInfo['price']['amount'];
        $quantity = $productInfo['quantity'];

        // Fetch categories information
        $categoriesNames = [];
        $categoriesIds = [];
        if (isset($productInfo['categories']) && is_array($productInfo['categories'])) {
            foreach ($productInfo['categories'] as $category) {
                $categoriesNames[] = $category['name'];
                $categoriesIds[] = $category['id'];
            }
        }

        $promotionTitle = $productInfo['promotion']['title'] ?? '';
        $metadataTitle = $productInfo['metadata']['title'] ?? '';
        $metadataDescription = $productInfo['metadata']['description'] ?? '';
        $description = $productInfo['description'];
        $productImageURLs = array_column($productInfo['images'], 'url');

        // Add product information to HTML table
        $htmlTable .= "<tr>
            <td>{$productId}</td>
            <td>{$productName}</td>
            <td>{$productPriceBeforeDiscount} {$productInfo['regular_price']['currency']}</td>
            <td>{$productPriceAfterDiscount} {$productInfo['price']['currency']}</td>
            <td>{$quantity}</td>
            <td>" . implode('<br>', $categoriesNames) . "</td>
            <td>" . implode('<br>', $categoriesIds) . "</td>
            <td>{$promotionTitle}</td>
            <td>{$metadataTitle}</td>
            <td>{$metadataDescription}</td>
            <td>{$description}</td>
            <td>" . implode('<br>', $productImageURLs) . "</td>
        </tr>";
    }

    // Close HTML table
    $htmlTable .= "</table>";

    // Export to HTML file
    file_put_contents('product_data.html', $htmlTable);

    // Export to Excel file
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();

    // Add data to Excel sheet header
    $sheet->setCellValue('A1', 'Product SKU');
    $sheet->setCellValue('B1', 'Product Name');
    $sheet->setCellValue('C1', 'Product Price Before Discount');
    $sheet->setCellValue('D1', 'Product Price After Discount');
    $sheet->setCellValue('E1', 'Quantity');
    $sheet->setCellValue('F1', 'Categories Names');
    $sheet->setCellValue('G1', 'Categories Ids');
    $sheet->setCellValue('H1', 'Promotion Title');
    $sheet->setCellValue('I1', 'Metadata Title');
    $sheet->setCellValue('J1', 'Metadata Description');
    $sheet->setCellValue('K1', 'Description');
    $sheet->setCellValue('L1', 'Product Image URLs');

    // Initialize row counter
    $row = 2;

    // Iterate over each product for Excel
    foreach ($jsonData['data'] as $productInfo) {
        // Extract and process product information
        $productId = $productInfo['id'];
        $productName = $productInfo['name'];
        $productPriceBeforeDiscount = $productInfo['regular_price']['amount'];
        $productPriceAfterDiscount = $productInfo['price']['amount'];
        $quantity = $productInfo['quantity'];

        // Fetch categories information
        $categoriesNames = [];
        $categoriesIds = [];
        if (isset($productInfo['categories']) && is_array($productInfo['categories'])) {
            foreach ($productInfo['categories'] as $category) {
                $categoriesNames[] = $category['name'];
                $categoriesIds[] = $category['id'];
            }
        }

        $promotionTitle = $productInfo['promotion']['title'] ?? '';
        $metadataTitle = $productInfo['metadata']['title'] ?? '';
        $metadataDescription = $productInfo['metadata']['description'] ?? '';
        $description = $productInfo['description'];
        $productImageURLs = array_column($productInfo['images'], 'url');

        // Add product information to Excel sheet
        $sheet->setCellValue('A' . $row, $productId);
        $sheet->setCellValue('B' . $row, $productName);
        $sheet->setCellValue('C' . $row, $productPriceBeforeDiscount);
        $sheet->setCellValue('D' . $row, $productPriceAfterDiscount);
        $sheet->setCellValue('E' . $row, $quantity);
        $sheet->setCellValue('F' . $row, implode("\n", $categoriesNames));
        $sheet->setCellValue('G' . $row, implode("\n", $categoriesIds));
        $sheet->setCellValue('H' . $row, $promotionTitle);
        $sheet->setCellValue('I' . $row, $metadataTitle);
        $sheet->setCellValue('J' . $row, $metadataDescription);
        $sheet->setCellValue('K' . $row, $description);
        $sheet->setCellValue('L' . $row, implode("\n", $productImageURLs));

        // Increment row counter
        $row++;
    }

    // Save Excel file
    $writer = new Xlsx($spreadsheet);
    $writer->save('product_data.xlsx');

    echo 'Exported HTML table and Excel file successfully.';
}
?>
