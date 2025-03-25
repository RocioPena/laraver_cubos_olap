<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Session;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ConsultaCluesController extends Controller
{
   
    public function mostrarFormulario()
    {
        return view('exportar');
    }

    
    public function consultarClues(Request $request)
    {
        $clues = $request->input('clues');
        $variables = array_map('trim', explode(',', $request->input('variables')));
        $modo = $request->input('modo', 'codigo+nombre');

      
        $response = Http::post('http://127.0.0.1:8070/consulta_avanzada', [
            "variables_clave" => $variables,
            "unidades" => ["[DIM UNIDAD].[CLUES].[{$clues}]"],
            "modo_encabezado" => $modo
        ]);

 
        if (!$response->ok()) {
            return back()->with('error', 'Error al consumir la API')->withInput();
        }

        $datos = $response->json();

        if (empty($datos)) {
            return back()->with('error', 'No se encontraron datos para los filtros seleccionados.')->withInput();
        }

   
        Session::put('datos_excel', $datos);

 
        return view('exportar', [
            'datos' => $datos,
            'clues' => $clues,
            'variables' => implode(',', $variables),
            'modo' => $modo
        ]);
    }


    public function descargarExcel()
    {
        $datos = Session::get('datos_excel');

        if (!$datos || count($datos) === 0) {
            return back()->with('error', 'No hay datos disponibles para exportar.');
        }


        


        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $sheet->fromArray(array_keys($datos[0]), null, 'A1');

        $row = 2;
        foreach ($datos as $item) {
            $sheet->fromArray(array_values($item), null, 'A' . $row++);
        }


        $filename = 'consulta_clues_' . now()->format('Ymd_His') . '.xlsx';
        $path = storage_path("app/public/{$filename}");

        $writer = new Xlsx($spreadsheet);
        $writer->save($path);

        return response()->download($path)->deleteFileAfterSend(true);
    }
}
