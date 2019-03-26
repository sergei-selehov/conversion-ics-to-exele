<?php

require_once 'vendor/autoload.php';
require_once 'src/ProgressBar/Manager.php';
require_once 'src/ProgressBar/Registry.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use ICal\ICal;

$firstFile = 'first/ICalFirst.ics';
$finishFile = 'finish/FileXlsx.xlsx';

if (file_exists($firstFile))
{
    unlink($firstFile);
}
if (file_exists($finishFile))
{
    unlink($finishFile);
}
foreach(glob('start/*.ics') as $file)
{
    file_put_contents('first/ICalFirst.ics',$file.' -> '.file_get_contents($file),FILE_APPEND);
}

try {
    $ical = new ICal('first/ICalFirst.ics', array(
        'defaultSpan'                 => 2,     // Default value
        'defaultTimeZone'             => 'UTC',
        'defaultWeekStart'            => 'MO',  // Default value
        'disableCharacterReplacement' => false, // Default value
        'filterDaysAfter'             => null,  // Default value
        'filterDaysBefore'            => null,  // Default value
        'replaceWindowsTimeZoneIds'   => false, // Default value
        'skipRecurrence'              => false, // Default value
        'useTimeZoneWithRRules'       => false, // Default value
    ));
} catch (\Exception $e) {
    die($e);
}

$forceTimeZone = false;

$showExample = array(
    'interval' => true,
    'range'    => true,
    'all'      => true,
);

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
if ($showExample['all']) {
    $events = $ical->sortEventsWithOrder($ical->events());
    $column = 1;
    $progressBar = new \ProgressBar\Manager(0, count($events));

    foreach ($events as $event){
        $column++;
        $dtstart = $ical->iCalDateToDateTime($event->dtstart, $forceTimeZone);
        if (isset($event->dtend)){
            $dtend = $ical->iCalDateToDateTime($event->dtend, $forceTimeZone);
            $dtend->format('d-m-Y H:i');
        } else {
            $dtend = 'NaN';
        }
        $dtstamp = $ical->iCalDateToDateTime($event->dtstamp, $forceTimeZone);
        $dtlast_modified = $ical->iCalDateToDateTime($event->last_modified, $forceTimeZone);

        $array_summary = explode("/", $event->summary);

        $sheet->setCellValue('A' . $column, $dtstart->format('d-m-Y H:i'));
        $sheet->setCellValue('B' . $column, $dtend);
        $sheet->setCellValue('C' . $column, $array_summary[0]);
        $sheet->setCellValue('D' . $column, $array_summary[1]);
        $sheet->setCellValue('E' . $column, $array_summary[2]);
        $sheet->setCellValue('F' . $column, $array_summary[3]);
        $sheet->setCellValue('G' . $column, $dtstamp->format('d-m-Y H:i'));
        $sheet->setCellValue('H' . $column, $event->uid);
        $sheet->setCellValue('I' . $column, $event->description);
        $sheet->setCellValue('J' . $column, $event->location);
        $sheet->setCellValue('K' . $column, $event->status);
        $sheet->setCellValue('L' . $column, $event->transp);
        $sheet->setCellValue('M' . $column, $event->sequence);
        $sheet->setCellValue('N' . $column, $dtlast_modified->format('d-m-Y H:i'));
//        echo $event->printData();
        $progressBar->advance();
    }
}

$writer = new Xlsx($spreadsheet);
$writer->save('finish/FileXlsx.xlsx');
echo 'Complete! Column: ' . $column . '; File: finish/FileXlsx.xlsx;';