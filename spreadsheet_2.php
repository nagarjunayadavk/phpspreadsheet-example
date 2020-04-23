<?php 
    
    echo "PHP has been installed successfully!";
    require 'vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    use PhpOffice\PhpSpreadsheet\Style\Fill;

    $jsonArray = '[
        {
            "title" => "Professional JavaScript",
            "author" => "Nicholas C. Zakas"
        },
        {
            "title" => "JavaScript: The Definitive Guide",
            "author" => "David Flanagan"
        },
        {
            "title" => "High Performance JavaScript",
            "author" => "Nicholas C. Zakas"
        }
    ]';

    $books = json_decode($jsonArray, true);
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    //headers
    $sheet->fromArray(array_keys($books[0]), NULL, 'A1');
    //getStyle accepts a range of cells as well!
    $lastColumn = $sheet->getHighestColumn();
    // echo "lastColumn ".$lastColumn;
    $sheet->getStyle('A1:'.$lastColumn.'1')->applyFromArray(
        array(
        'fill' => array(
            'type' => Fill::FILL_SOLID,
            'color' => array('rgb' => 'E5E4E2' )
        ),
        'font'  => array(
            'bold'  =>  true
        )
        )
    );
    // echo "naga --".$books[1]['TestCase ID'].'---';
    $x = 0;
    while($x <= count($books)) {
        echo "Title is: ".$books[$x]['title']." <br>";
        $col = 'A';
        foreach(array_keys($books[0]) as $value) {
            echo "The Title is: ".$books[$x][$value]." <br>";
            $sheet->setCellValue($col++.($x+2), $books[$x][$value]);
        }
        $x++;
    }

    $spreadsheet->getActiveSheet()->setTitle('BooksList');
    $writer = new Xlsx($spreadsheet);
    $writer->save('BooksList.xlsx');

    echo "<meta http-equiv='refresh' content='0;url=BooksList.xlsx'/>";