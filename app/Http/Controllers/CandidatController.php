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


class CandidatController extends Controller
{   
    public function exportToExcel(){
        //On recuper les données depuis le model
        $candidats= Candidat::with('activite')->get();
        //on cree un nouveau classeur PhpSpreadsheet
        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

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

        $worksheet->mergeCells('A2:M2');
        $worksheet->setCellValue('A2', 'learner data');
        $worksheet->getStyle('A2')
        ->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->setStartColor(new Color('FF87CEEB'));
    
    $worksheet->getStyle('A2')
        ->getAlignment()
        ->setHorizontal(Alignment::HORIZONTAL_CENTER);


// Définir le style de la bordure
funtion applyColorToColumn(Worksheet $worksheet, int $startRow, int $maxRows, string $columnLetter, string $color)
{ //la couleur aux lignes spécifiques

    for ($row = $startRow; $row<= $startRow + $maxRows -1; $row++){
        $worksheet->getStyle($columnLetter . $row)->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->getStartColor()->setARGB();
    }

    //couleur pour les lignes suivante

    $lastRow =$worksheet->getHighestColumn();
    for($row = $startRow +$maxRows; $row <= $lastRow; $row++){
        $worksheet->getStyle($columnletter . $row)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB($color);
    }
}

$worksheet=$spreadsheet->getActiveSheet();
applyColorToColumn($worksheet, 3, 50, 'B', 'FFFFE0');


// Définir la couleur pour les lignes suivantes si nécessaire
    

        //on  defini les en-têtes de colonne
        $worksheet->setCellValue('A3', 'Id')
                    ->getStyle('A3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('B3', 'name')
                    ->getStyle('B3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('C3', 'last_name')
                    ->getStyle('C3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('D3', 'gender')
                    ->getStyle('D3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);    
        $worksheet->setCellValue('E3', 'Years')
                    ->getStyle('E3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('F3', 'socio_professional')
                    ->getStyle('F3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('G3', 'student_university')
                    ->getStyle('G3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('H3', 'student_speciality')
                    ->getStyle('H3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('I3', 'e_mail')
                    ->getStyle('I3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('J3', 'phone_number')
                    ->getStyle('J3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('K3', 'linkedin')
                    ->getStyle('K3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);
        $worksheet->setCellValue('L3', 'nome activités')
                    ->getStyle('L3')
                    ->getFont()
                    ->setSize(12)
                    ->setBold(true);


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
