<?php
/**This code fetches the product list from the Salla API,
 *  processes the data, and generates an HTML table and an Excel file for each product. 
 * Make sure to replace placeholders like 'your clientId', 'your clientSecret',
 *  and 'your refreshToken' with your actual Salla API credentials. */


// Include the necessary libraries
require_once 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Get access token from environment variable
$accessToken = getenv('Access_TOKEN');

// Salla API endpoint to get the product list
$apiEndpoint = "https://api.salla.dev/admin/v2/products";

// Initialize cURL
$curl = curl_init();

// Set cURL options
curl_setopt_array($curl, [
    CURLOPT_URL => $apiEndpoint,
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_ENCODING => "",
    CURLOPT_MAXREDIRS => 10,
    CURLOPT_TIMEOUT => 30,
    CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
    CURLOPT_CUSTOMREQUEST => "GET",
    CURLOPT_HTTPHEADER => [
        "Accept: application/json",
        "Authorization: Bearer " . $accessToken
    ],
]);

// Execute cURL and get the response
$response = curl_exec($curl);
$err = curl_error($curl);

// Close cURL
curl_close($curl);

// Check for cURL errors
if ($err) {
    echo "cURL Error #:" . $err;
} else {
    // Decode the JSON data
    $jsonData = json_decode($response, true);

    // Check if decoding was successful and if the data structure is valid
    if ($jsonData === null || !isset($jsonData['data']) || !is_array($jsonData['data'])) {
        echo 'Error decoding JSON response or invalid data structure.';
    } else {
        // Process product data
        $htmlTable = generateHtmlTable($jsonData['data']);
        $excelFilePath = generateExcelFile($jsonData['data']);

        // Create a ZIP archive containing the HTML and Excel files
        $zipFilePath = createZipArchive($htmlTable, $excelFilePath);
        if ($zipFilePath) {
            // Prompt the user to download the ZIP archive
            header('Content-Type: application/zip');
            header('Content-Disposition: attachment; filename="product_data.zip"');
            header('Content-Length: ' . filesize($zipFilePath));
            readfile($zipFilePath);

            // Clean up temporary files
            unlink($zipFilePath);
            unlink($htmlFilePath);
            unlink($excelFilePath);
        } else {
            echo "Failed to create ZIP file";
        }
    }
}

function generateHtmlTable($data) {
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
    foreach ($data as $productInfo) {
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
                <td>{$productPriceBeforeDiscount}</td>
                <td>{$productPriceAfterDiscount}</td>
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
    $htmlFilePath = '/tmp/product_data.html';
    file_put_contents($htmlFilePath, $htmlTable);

    return $htmlTable;
}

function generateExcelFile($data) {
    // Export to Excel file
    $excelFilePath = '/tmp/product_data.xlsx';
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
    foreach ($data as $productInfo) {
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
    $writer->save($excelFilePath);

    return $excelFilePath;
}

function createZipArchive($htmlTable, $excelFilePath) {
    $zip = new ZipArchive();
    $zipFileName = "/tmp/product_data.zip";
    if ($zip->open($zipFileName, ZipArchive::CREATE) === TRUE) {
        $zip->addFromString('product_data.html', $htmlTable);
        $zip->addFile($excelFilePath, 'product_data.xlsx');
        $zip->close();
        return $zipFileName;
    }
    return null;
}
?>
