<?php 
    
    echo "PHP has been installed successfully!";
    require 'vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
    use PhpOffice\PhpSpreadsheet\Style\Fill;
    $books = array(
        array(
            "title" => "Professional JavaScript",
            "author" => "Nicholas C. Zakas"
        ),
        array(
            "title" => "JavaScript: The Definitive Guide",
            "author" => "David Flanagan"
        ),
        array(
            "title" => "High Performance JavaScript",
            "author" => "Nicholas C. Zakas"
        )
    );
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    //headers
    $sheet->fromArray(array_keys($books[0]), NULL, 'A1');
    //getStyle accepts a range of cells as well!
    $sheet->getStyle('A1:B1')->applyFromArray(
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
    $x = 0;
    while($x <= count($books)) {
        // echo "The number is: ".$books[$x]['title']." <br>";
        $sheet->setCellValue('A'.($x+2), $books[$x]['title']);
        $sheet->setCellValue('B'.($x+2), $books[$x]['author']);
        $x++;
    }

    $spreadsheet->getActiveSheet()->setTitle('Books');

    $writer = new Xlsx($spreadsheet);
    $writer->save('test.xlsx');

    echo "<meta http-equiv='refresh' content='0;url=test.xlsx'/>";