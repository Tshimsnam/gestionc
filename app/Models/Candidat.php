<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\BelongsTo;
use Illuminate\Database\Eloquent\Factories\HasFactory;

class Candidat extends Model
{
    use HasFactory;
    protected $guarded=[];

    public function activite() : BelongsTo{
        return $this->belongsTo(Activite::class);
    }

    public function stage(){
        return $this->belongsTo(Stage::class);
    }
}
