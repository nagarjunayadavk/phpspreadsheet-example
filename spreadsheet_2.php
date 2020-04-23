<?php 
    
    echo "PHP has been installed successfully!";
    require 'vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    use PhpOffice\PhpSpreadsheet\Style\Fill;

    $jsonArray = '[
        {
            "TestCase ID": "1",
            "Test Input": "Challenging Times",
            "Expected Output": "Business",
            "Actulal OutPut": "125.60",
            "Result": "Pass"
        },
        {
            "TestCase ID": "2",
            "Test Input": "Learning JavaScript",
            "Expected Output": "Programming",
            "Actulal OutPut": "56.00",
            "Result": "Pass"
        },
        {
            "TestCase ID": "3",
            "Test Input": "Popular Science",
            "Expected Output": "Science",
            "Actulal OutPut": "210.40",
            "Result": "Pass"
        }
    ]';

    $books = json_decode($jsonArray, true);
    // $books = array(
    //     array(
    //         "title" => "Professional JavaScript",
    //         "author" => "Nicholas C. Zakas"
    //     ),
    //     array(
    //         "title" => "JavaScript: The Definitive Guide",
    //         "author" => "David Flanagan"
    //     ),
    //     array(
    //         "title" => "High Performance JavaScript",
    //         "author" => "Nicholas C. Zakas"
    //     )
    // );
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
        echo "The number is: ".$books[$x]['TestCase ID']." <br>";
        $col = 'A';
        foreach(array_keys($books[0]) as $value) {
            echo "The number is: ".$books[$x][$value]." <br>";
            $sheet->setCellValue($col++.($x+2), $books[$x][$value]);
        }
        // $sheet->setCellValue('A'.($x+2), $books[$x]['TestCase ID']);
        // $sheet->setCellValue('B'.($x+2), $books[$x]['Test Input']);
        // $sheet->setCellValue('C'.($x+2), $books[$x]['Expected Output']);
        // $sheet->setCellValue('D'.($x+2), $books[$x]['Actulal OutPut']);
        // $sheet->setCellValue('E'.($x+2), $books[$x]['Result']);
        $x++;
    }

    $spreadsheet->getActiveSheet()->setTitle('TestCasesList');
    $writer = new Xlsx($spreadsheet);
    $writer->save('TestCasesList.xlsx');

    echo "<meta http-equiv='refresh' content='0;url=TestCasesList.xlsx'/>";