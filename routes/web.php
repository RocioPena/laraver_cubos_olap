<?php

use App\Http\Controllers\ConsultaCluesController;

Route::get('/', [ConsultaCluesController::class, 'mostrarFormulario'])->name('formulario.clues');
Route::post('/consultar-clues', [ConsultaCluesController::class, 'consultarClues'])->name('consultar.clues');
Route::get('/descargar-excel', [ConsultaCluesController::class, 'descargarExcel'])->name('descargar.excel');





