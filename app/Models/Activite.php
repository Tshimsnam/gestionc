<?php

namespace App\Models;

use App\Models\Candidat;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Factories\HasFactory;

class Activite extends Model
{
    use HasFactory;
    protected $fillable = ['name', 'place', 'duree','date_debut'];

    public function candidats(){
        return $this->hasMany(Candidat::class);
    }
}
