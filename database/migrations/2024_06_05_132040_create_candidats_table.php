<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;
use Ramsey\Uuid\Type\Integer;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
        Schema::create('candidats', function (Blueprint $table) {
            $table->id();
            $table->string('name');
            $table->string('last_name');
            $table->string('gender');
            $table->string('years');
            $table->string('socio_professional');
            $table->string('student_university');
            $table->string('student_speciality');
            $table->string('e_mail');
            $table->string('phone_number');
            $table->string('linkedin');
            $table->foreignId('activites_id')->constrained()->onDelete('cascade')->onUpdate('cascade');
            $table->foreignId('stages_id')->constrained()->onDelete('cascade')->onUpdate('cascade');
            $table->timestamps();
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('candidats');
    }
};
