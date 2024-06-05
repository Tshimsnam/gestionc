<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class Candidat extends Model
{
    use HasFactory;

    public function activites(){
        return $this->hasMany(Activite::class);
    }

    public function stages(){
        return $this->hasMany(Stage::class);
    }
}
