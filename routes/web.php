<?php

use App\Http\Controllers\CandidatController;
use App\Http\Controllers\ProfileController;
use App\Models\Candidat;
use Illuminate\Support\Facades\Route;

Route::get('/', function () {
    return view('welcome');
});

Route::get('/dashboard', function () {
    return view('dashboard');
})->middleware(['auth', 'verified'])->name('dashboard');

Route::middleware('auth')->group(function () {
    Route::get('/profile', [ProfileController::class, 'edit'])->name('profile.edit');
    Route::patch('/profile', [ProfileController::class, 'update'])->name('profile.update');
    Route::delete('/profile', [ProfileController::class, 'destroy'])->name('profile.destroy');
    Route::get('/candidats', [CandidatController::class, 'index'])->name('candidat.index');
    Route::get('/exportToExcel',[CandidatController::class, 'exportToExcel'])->name('candidats.exportToExcel');
});

require __DIR__.'/auth.php';
