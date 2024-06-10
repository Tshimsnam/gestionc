<?php

namespace App\Http\Controllers;

use App\Models\Stage;
use App\Models\Candidat;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use App\Http\Requests\StoreCandidatRequest;
use App\Http\Requests\UpdateCandidatRequest;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;


class CandidatController extends Controller
{   
    public function exportToExcel(){
        //On recuper les données depuis le model
        $candidats= Candidat::with('activite')->get();
        //on cree un nouveau classeur PhpSpreadsheet
        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        //premiere section du tableau
        $worksheet->mergeCells('A1:M1');
        $worksheet->setCellValue('A1', 'Data');
        $worksheet->getStyle('A1')
        ->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->setStartColor(new Color('FFFFFF00'));
    //pour l'allignement
    $worksheet->getStyle('A1')
        ->getAlignment()
        ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $worksheet->mergeCells('A2:L2');
        $worksheet->setCellValue('A2', 'learner data');
        $worksheet->getStyle('A2')
        ->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->setStartColor(new Color('FF87CEEB'));
    
    $worksheet->getStyle('A2')
        ->getAlignment()
        ->setHorizontal(Alignment::HORIZONTAL_CENTER);


// fonction pour les couleurs des colonnes
function applyColorToColumn(Worksheet $worksheet, int $startRow, int $maxRows, string $columnLetter, string $color)
{ //la couleur aux lignes spécifiques

    for ($row = $startRow; $row<= $startRow + $maxRows -1; $row++){
        $worksheet->getStyle($columnLetter . $row)->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->getStartColor()->setARGB($color);
    }

    //couleur pour les lignes suivante

    $lastRow =$worksheet->getHighestRow();
    for($row = $startRow +$maxRows; $row <= $lastRow; $row++){
        $worksheet->getStyle($columnLetter . $row)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB($color);
    }
}

//Colonne colorer
$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 4, 50, 'D', 'FFFFE0');

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 4, 50, 'E', 'FFFFE0');

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 4, 50, 'F', 'FFFFE0');

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 4, 50, 'N', 'FFFFE0' );

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 4, 50, 'S', 'FFFFE0');

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 4, 50, 'Z', 'FFFFE0');

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 4, 50, 'AA', 'FFFFE0');

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 4, 50, 'AF', 'FFFFE0');

//colonne separateur de section

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 1, 50, 'M', 'FF808080');

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 1, 50, 'R', 'FF808080');

$worksheet=$spreadsheet->getActiveSheet(); 
applyColorToColumn($worksheet, 1, 50, 'Y', 'FF808080');

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 1, 50, 'AD', 'FF808080');

    
        //on  defini les en-têtes de colonne
        $worksheet->setCellValue('A3', 'Id')
                    ->getStyle('A3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);        
        $worksheet->setCellValue('B3', 'name')
                    ->getStyle('B3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('C3', 'last_name')
                    ->getStyle('C3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('D3', 'gender')
                    ->getStyle('D3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);    
        $worksheet->setCellValue('E3', 'Years')
                    ->getStyle('E3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('F3', 'socio_professional')
                    ->getStyle('F3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('G3', 'student_university')
                    ->getStyle('G3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('H3', 'student_speciality')
                    ->getStyle('H3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('I3', 'e_mail')
                    ->getStyle('I3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('J3', 'phone_number')
                    ->getStyle('J3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('K3', 'linkedin')
                    ->getStyle('K3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('L3', 'nome activités')
                    ->getStyle('L3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);

        // Centrer les en-têtes
$worksheet->getStyle('A3:AH3')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

// Descendre les titres longs sur plusieurs lignes
$worksheet->getCell('F3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('G3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('H3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('O3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('P3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('Q3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('R3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('S3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('T3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('U3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('V3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('W3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('X3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('AB3')->getStyle()->getAlignment()->setWrapText(true);
$worksheet->getCell('AC3')->getStyle()->getAlignment()->setWrapText(true);


        
        //la taille (longueur) des lignes des entetes

        $worksheet->getRowDimension(3)->setRowHeight(25);
        //longueur des lignes
        $startColumn = 'A'; // Première colonne à modifier
        $endColumn = 'AH'; // Dernière colonne à modifier
        $newColumnWidth = 25; // Nouvelle largeur de colonne en points

for ($column = $startColumn; $column <= $endColumn; $column++) {
    $worksheet->getColumnDimension($column)->setWidth($newColumnWidth);
}

//ligne de separation et taille
$columnLetter = 'M'; // Colonne à modifier
$newColumnWidth = 4; // Nouvelle largeur de colonne en points

$worksheet->getColumnDimension($columnLetter)->setWidth($newColumnWidth);

$columnLetter = 'R';
$newColumnWidth = 4;

$worksheet->getColumnDimension($columnLetter)->setWidth($newColumnWidth);

$columnLetter = 'Y';
$newColumnWidth = 4;
$worksheet->getColumnDimension($columnLetter)->setWidth($newColumnWidth);

$columnLetter = 'AD';
$newColumnWidth = 4;
$worksheet->getColumnDimension($columnLetter)->setWidth($newColumnWidth);   


//Remplissage de données 
        $row = 4;

        foreach($candidats as $candidat){
            $worksheet->setCellValue('A'.$row, $candidat->id)
                        ->getStyle('A'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('B'.$row, $candidat->name)
                        ->getStyle('B'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('C'.$row, $candidat->last_name)
                        ->getStyle('C'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('D'.$row, $candidat->gender)
                        ->getStyle('D'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('E'.$row, $candidat->years)
                        ->getStyle('E'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('F'.$row, $candidat->socio_professional)
                        ->getStyle('F'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('G'.$row, $candidat->student_university)
                        ->getStyle('G'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('H'.$row, $candidat->student_speciality)
                        ->getStyle('H'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('I'.$row, $candidat->e_mail)
                        ->getStyle('I'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('J'.$row, $candidat->phone_number)
                        ->getStyle('J'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('K'.$row, $candidat->linkedin)
                        ->getStyle('K'.$row)
                        ->getAlignment()
                        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
            $worksheet->setCellValue('L'.$row, $candidat->activite->name);

                        $worksheet->getRowDimension($row)->setRowHeight(-1);
            $row++;
        } 


        //Deuxieme section du tableau

        $worksheet->mergeCells('N1:AD1');
        $worksheet->setCellValue('N1', 'Participation in : Training/Internship/Event');
        $worksheet->getStyle('N1')
                    ->getFill()
                    ->setFillType(FILL::FILL_SOLID)
                    ->setStartColor(new Color('FFF80F00'));

        //allignement de mergecellule

        $worksheet->getStyle('N1')
                    ->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        //sous section
        $worksheet->mergeCells('N2:Q2');
        $worksheet->setCellValue('N2', 'To be specified if Training');
        $worksheet->getStyle('N2')
                    ->getFill()
                    ->setFillType(FILL::FILL_SOLID)
                    ->setStartColor(new Color('FFCC9966'));

        $worksheet->getStyle('N2')
                    ->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $worksheet->setCellValue('N3', 'Place')
                    ->getStyle('N3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('O3', 'Training Date')
                    ->getStyle('O3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('P3', 'Training Duration')
                    ->getStyle('P3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('Q3', 'Name of Training Conducted')
                    ->getStyle('Q3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);

    // remplissage des données 
        //.....


          //3ème section du tableau

        $worksheet->mergeCells('S2:X2');
        $worksheet->setCellValue('S2', 'To be specified if internship');
        $worksheet->getStyle('S2')
                    ->getFill()
                    ->setFillType(FILL::FILL_SOLID)
                    ->setStartColor(new Color('FFDD8822'));
  
        $worksheet->getStyle('S2')
                    ->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $worksheet->setCellValue('S3', 'Place')
                    ->getStyle('S3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('T3', 'Internship Start date')
                    ->getStyle('T3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('U3', 'Internship end date')
                    ->getStyle('U3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('V3', 'Internship duration (number of days')
                    ->getStyle('V3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('W3', 'Intership Name')
                    ->getStyle('W3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('X3', 'Project/Solutions developed')
                    ->getStyle('X3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);

        //remplissage des données





        //4ème section du tableau
        $worksheet->mergeCells('Z2:AC2');
        $worksheet->setCellValue('Z2', 'To be specified if Event');
        $worksheet->getStyle('Z2')
                    ->getFill()
                    ->setFillType(FILL::FILL_SOLID)
                    ->setStartColor(new Color('FFFFC0CB'));
        
        $worksheet->getStyle('Z2')
                    ->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        
        $worksheet->setCellValue('Z3', 'Place')
                    ->getStyle('Z3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('AA3', 'Event date')
                    ->getStyle('AA3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('AB3', 'Event duration(number od days)')
                    ->getStyle('AB3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('AC3', 'Name of the Event')
                    ->getStyle('AC3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);

        //remplissage des donnes



        //4ème section du tableau

        $worksheet->mergeCells('AE1:AG1');
        $worksheet->setCellValue('AE1', 'Jobs Created');
        $worksheet->getStyle('AE1')
                    ->getFill()
                    ->setFillType(FILL::FILL_SOLID)
                    ->setStartColor(new Color('FF00FF00'));

        $worksheet->getStyle('AE1')
                    ->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        // Élargir la largeur des colonnes de AE1 à AG1
        $startColumn = 'AE';
        $endColumn = 'AG';
        $columnWidth = 15; // Définissez la largeur souhaitée en points
        for ($column = $startColumn; $column <= $endColumn; $column++) {
            $worksheet->getColumnDimension($column)->setWidth($columnWidth);
        }
        

        $worksheet->mergeCells('AE2:AG2');
        $worksheet->setCellValue('AE2', 'Employement of beneficiaires after ODC');
        $worksheet->getStyle('AE2')
                    ->getFill()
                    ->setFillType(FILL::FILL_SOLID)
                                ->setStartColor(new Color('FFB0E0E6'));

        $worksheet->getStyle('AE2')
                    ->getAlignment()
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $worksheet->setCellValue('AE3', 'Company Name')
                    ->getStyle('AE3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('AF3', 'Contract type')
                    ->getStyle('AF3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('AG3', 'Job Position')
                    ->getStyle('AG3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);
        $worksheet->setCellValue('AH3', 'Comments')
                    ->getStyle('AH3')
                    ->getFont()
                    ->setSize(10)
                    ->setBold(true);


        /*foreach ($worksheet->getColumnIterator() as $column) {
            $worksheet->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
        }
        $worksheet->getStyle('A1:K1')->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('0070C0');*/
        
  

        //on cree le fichier Excel
        $writer = new Xlsx($spreadsheet); 
        $fileName ='candidats.Xlsx';
        $tempFile = tempnam(sys_get_temp_dir(), $fileName); 
        $writer->save($tempFile);

        //On renvoie le fichier Excel au navigateur

        return response()->download($tempFile, $fileName)->deleteFileAfterSend(true);
    }


    /**
     * Display a listing of the resource.
     */
    public function index()
    {
        return view('rapportexcel.index');
    }

    /**
     * Show the form for creating a new resource.
     */
    public function create()
    {
        //
    }

    /**
     * Store a newly created resource in storage.
     */
    public function store(StoreCandidatRequest $request)
    {
        //
    }

    /**
     * Display the specified resource.
     */
    public function show(Candidat $candidat)
    {
        //
    }

    /**
     * Show the form for editing the specified resource.
     */
    public function edit(Candidat $candidat)
    {
        //
    }

    /**
     * Update the specified resource in storage.
     */
    public function update(UpdateCandidatRequest $request, Candidat $candidat)
    {
        //
    }

    /**
     * Remove the specified resource from storage.
     */
    public function destroy(Candidat $candidat)
    {
        //
    }
}
