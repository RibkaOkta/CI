<?php

namespace App\Controllers;

use App\Controllers\BaseController;
use App\Models\UserModel;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Import extends BaseController
{
    public function __construct()
    {
        $this->user = new UserModel();
        // helper('form'); //atau bisa tambahkan di BaseController dibagian protected $helpers = ['form'];
    }

    public function index()
    {
        $data = [
            'users' => $this->user->findAll()
        ];

        return view("view_import", $data);
    }

    public function upload()
    {
        if (!$this->validate([
            'excel' => [
                'rules' => 'uploaded[excel]|max_size[excel,10240]|ext_in[excel,xls,xlsx]',
                'errors' => [
                    'uploaded' => 'File Upload Masih Kosong',
                    'max_size' => 'Max ukuran file 10 Mb',
                    'ext_in' => 'File Harus dalam bentuk .xls atau .xlsx',
                ]
            ],
        ])) {
            session()->setFlashdata('error', $this->validator->listErrors());
            return redirect()->back()->withInput();
        }

        $file_excel = $this->request->getFile('excel');

        $extensi = $file_excel->getClientExtension();

        if ($extensi == 'xls') {
            $render = new \PhpOffice\PhpSpreadsheet\Reader\Xls();
        } else if ($extensi == 'xlsx') {
            $render = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        } else {
            session()->setFlashdata('error', 'Error');
            return redirect()->back()->withInput();
        }

        $preadsheet = $render->load($file_excel);
        $data = $preadsheet->getActiveSheet()->toArray();

        $dataGagal = 0;
        $dataBerhasil = 0;
   //   dd($data);
        foreach ($data as $key => $row) {
            if ($key < 4 or $row[2] == null) {
                continue;
            }

            $nama = $row[1];
            $nip = $row[2];
            $unit_kerja = $row[3];
            $kabkota = $row[4];
            $email = $row[5];
            $no_hp = $row[6];



            $datauser = $this->user->where('nip', $nip)->findAll();
            if ($datauser) {
                $dataGagal++;
            } else {
            
                $user = [
                    'nama' => $nama,
                    'nip' => $nip,
                    'unit_kerja' => $unit_kerja,
                    'kabkota' => $kabkota,
                    'email' => $email,
                    'no_hp' => $no_hp,
                ];
                $this->user->insert($user);
                $dataBerhasil++;
            }
        }

        //Jika file ingin di simpan setiap upload
        // if ($file_excel->isValid() && !$file_excel->hasMoved()) {
        //     $excelName = $file_excel->getRandomName();
        //     $file_excel->move('file_excel/', $excelName);
        // }

        return redirect()->to('Import/index')->with('sukses', $dataGagal . ' Data Duplikat, Gagal Ditambahkan <br>' . $dataBerhasil . ' Data Berhasil Ditambahkan <br>');
    }

    function download()
    {
        return $this->response->download('file_testing/Data_Mahasiswa.xlsx', null);
    }

    public function delete($id)
    {
        $this->user->delete($id);
        return redirect()->to('Import/index')->with('sukses', 'Data Berhasil Dihapus');
    }

    public function export()
    {

        $mahasiswa = $this->user->findAll();

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $spreadsheet->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Nim')
            ->setCellValue('B1', 'Nama');

        $column = 2;

        foreach ($mahasiswa as $mhs) {
            $spreadsheet->setActiveSheetIndex(0)
                ->setCellValue('A' . $column, $mhs['nim'])
                ->setCellValue('B' . $column, $mhs['nama']);

            $column++;
        }

        $styleArray = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
        $i = $column - 1;
        $sheet->getStyle('A1:B' . $i)->applyFromArray($styleArray);
        $sheet->getColumnDimension('A')->setAutoSize(TRUE);

        $writer = new Xlsx($spreadsheet);
        $filename = date('Y-m-d-His') . '-Data-Mahasiswa';

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename=' . $filename . '.xlsx');
        header('Cache-Control: max-age=0');

        $writer->save('php://output');
        header("Content-Type: application/vnd.ms-excel");
        exit;
    }
}
