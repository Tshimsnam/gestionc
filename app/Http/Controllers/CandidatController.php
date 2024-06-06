<?php

namespace App\Http\Controllers;

use App\Models\Stage;
use App\Models\Candidat;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use App\Http\Requests\StoreCandidatRequest;
use App\Http\Requests\UpdateCandidatRequest;


class CandidatController extends Controller
{   
    public function exportToExcel(){
        //On recuper les données depuis le model
        $candidats= Candidat::with('activite')->get();
        //on cree un nouveau classeur PhpSpreadsheet
        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();

        //on  defini les en-têtes de colonne
        $worksheet->setCellValue('A1', 'Id')
                    ->getStyle('A1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('B1', 'name')
                    ->getStyle('B1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('C1', 'last_name')
                    ->getStyle('C1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('D1', 'gender')
                    ->getStyle('D1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);    
        $worksheet->setCellValue('E1', 'Years')
                    ->getStyle('E1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('F1', 'socio_professional')
                    ->getStyle('F1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('G1', 'student_university')
                    ->getStyle('G1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('H1', 'student_speciality')
                    ->getStyle('H1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('I1', 'e_mail')
                    ->getStyle('I1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('J1', 'phone_number')
                    ->getStyle('J1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('K1', 'linkedin')
                    ->getStyle('K1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);
        $worksheet->setCellValue('L1', 'nome activités')
                    ->getStyle('L1')
                    ->getFont()
                    ->setSize(14)
                    ->setBold(true);


        //Remplissage de données 
        $row = 2;

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

        foreach ($worksheet->getColumnIterator() as $column) {
            $worksheet->getColumnDimension($column->getColumnIndex())->setAutoSize(true);
        }
        $worksheet->getStyle('A1:K1')->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('0070C0');
        
  

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
