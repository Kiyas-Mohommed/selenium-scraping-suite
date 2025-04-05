<?php

namespace Facebook\WebDriver;

set_time_limit(0);

use Exception;
use Facebook\WebDriver\WebDriverBy;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Facebook\WebDriver\Chrome\ChromeOptions;
use Facebook\WebDriver\Remote\RemoteWebDriver;
use Facebook\WebDriver\Remote\DesiredCapabilities;
use Facebook\WebDriver\WebDriverExpectedCondition;

require_once __DIR__ . '/vendor/autoload.php';

$serverUrl = 'http://localhost:4444/';
$capabilities = DesiredCapabilities::chrome();
$chromeOptions = new ChromeOptions();
$chromeOptions->addArguments(['--window-size=1024,768']);
$capabilities->setCapability(ChromeOptions::CAPABILITY, $chromeOptions);
$driver = RemoteWebDriver::create($serverUrl, $capabilities);

$progressFile = 'progress.txt';
$startPage = 1;

if (file_exists($progressFile)) {
    $startPage = (int) file_get_contents($progressFile);
    echo "Resuming from page: $startPage" . "\n";
}

$pagesPerBatch = 2000;
$currentBatch = intdiv($startPage - 1, $pagesPerBatch) + 1;

function getBatchFileName($batchNumber)
{
    return "scraped_data_batch_{$batchNumber}.xlsx";
}

function loadOrCreateSpreadsheet($batchNumber)
{
    $spreadsheetFile = getBatchFileName($batchNumber);

    if (file_exists($spreadsheetFile)) {
        echo "Loading existing batch file: $spreadsheetFile" . "\n";
        $spreadsheet = IOFactory::load($spreadsheetFile);
        $sheet = $spreadsheet->getActiveSheet();
        $rowNum = $sheet->getHighestRow() + 1;
    } else {
        echo "Creating new batch file: $spreadsheetFile" . "\n";
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Scraped Data');
        $sheet->setCellValue('A1', 'Part No');
        $sheet->setCellValue('B1', 'Description');
        $sheet->setCellValue('C1', 'Quantity');
        $rowNum = 2;
    }

    return [$spreadsheet, $sheet, $rowNum];
}

list($spreadsheet, $sheet, $rowNum) = loadOrCreateSpreadsheet($currentBatch);
$uniquePartNos = [];

function saveSpreadsheet($spreadsheet, $currentBatch)
{
    $writer = new Xlsx($spreadsheet);
    $writer->save(getBatchFileName($currentBatch));
}

function logTime($label)
{
    echo "[" . date('Y-m-d H:i:s') . "] $label" . "\n";
}

try {
    logTime("Navigating to the initial page $startPage...");
    $startTime = microtime(true);
    $driver->get("https://catalog.locatory.com/BrooksandMaldiniCorporation?page=$startPage");
    logTime("Page navigation completed.");
    echo "Time taken for page navigation: " . (microtime(true) - $startTime) . " seconds\n";

    try {
        logTime("Accepting cookie consent...");
        $startTime = microtime(true);
        $driver->findElement(WebDriverBy::id('CybotCookiebotDialogBodyLevelButtonLevelOptinAllowallSelection'))->click();
        echo "Time taken to handle cookies: " . (microtime(true) - $startTime) . " seconds\n";
    } catch (Exception $e) {
        echo "Cookie consent already handled or not required.\n";
    }

    logTime("Fetching total number of pages...");
    $startTime = microtime(true);
    $total_pages_text = $driver->findElement(WebDriverBy::cssSelector('#topPaging > div:nth-child(1) > span'))->getText();
    $total_pages = intval($total_pages_text);
    logTime("Total pages fetched.");
    echo "Time taken to fetch total pages: " . (microtime(true) - $startTime) . " seconds\n";
    echo "Total pages: $total_pages" . "\n";

    for ($page = $startPage; $page <= $total_pages; $page++) {
        $pageStartTime = microtime(true);
        logTime("Starting page $page");

        if (($page - 1) % $pagesPerBatch == 0 && $page > $startPage) {
            saveSpreadsheet($spreadsheet, $currentBatch);
            $currentBatch++;
            list($spreadsheet, $sheet, $rowNum) = loadOrCreateSpreadsheet($currentBatch);
        }

        $url = "https://catalog.locatory.com/BrooksandMaldiniCorporation?page=$page";
        logTime("Navigating to URL: $url");
        $startTime = microtime(true);
        $driver->get($url);
        echo "Time taken for page $page navigation: " . (microtime(true) - $startTime) . " seconds\n";

        logTime("Waiting for table to load...");
        $startTime = microtime(true);
        $driver->wait(2, 200)->until(
            WebDriverExpectedCondition::presenceOfElementLocated(WebDriverBy::id('sidebarLogo'))
        );
        echo "Time taken to wait for the table on page $page: " . (microtime(true) - $startTime) . " seconds\n";

        logTime("Extracting data from the page...");
        $startTime = microtime(true);
        $tbodyElements = $driver->findElements(WebDriverBy::tagName('table'));
        foreach ($tbodyElements as $tbody) {
            $rows = $tbody->findElements(WebDriverBy::tagName('tr'));

            foreach ($rows as $row) {
                $cells = $row->findElements(WebDriverBy::tagName('td'));
                $rowData = [];

                foreach ($cells as $cell) {
                    $rowData[] = $cell->getText();
                }
                echo "Row Data: " . implode(", ", $rowData) . "\n";

                if (count($rowData) >= 4) {
                    $partNo = $rowData[0];

                    if (!isset($uniquePartNos[$partNo])) {
                        $sheet->setCellValue('A' . $rowNum, $partNo);
                        $sheet->setCellValue('B' . $rowNum, $rowData[1]);
                        $sheet->setCellValue('C' . $rowNum, $rowData[3]);
                        $rowNum++;

                        $uniquePartNos[$partNo] = true;
                    }
                }
            }
        }
        echo "Time taken to extract data from page $page: " . (microtime(true) - $startTime) . " seconds\n";

        logTime("Saving progress...");
        $startTime = microtime(true);
        file_put_contents($progressFile, $page);
        saveSpreadsheet($spreadsheet, $currentBatch);
        echo "Time taken to save data for page $page: " . (microtime(true) - $startTime) . " seconds\n";

        echo "Total time for page $page: " . (microtime(true) - $pageStartTime) . " seconds\n";
    }
} catch (Exception $e) {
    echo "An error occurred: " . $e->getMessage() . "\n";
}

echo "Scraping complete. Data saved to multiple batch files." . "\n";
$driver->quit();
