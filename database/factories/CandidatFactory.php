<?php

namespace Database\Factories;

use Illuminate\Database\Eloquent\Factories\Factory;

/**
 * @extends \Illuminate\Database\Eloquent\Factories\Factory<\App\Models\Candidat>
 */
class CandidatFactory extends Factory
{
    /**
     * Define the model's default state.
     *
     * @return array<string, mixed>
     */
    public function definition(): array
    {
        return [
            'name'=>$this->faker->name,
            'last_name'=>$this->faker->lastName,
            'gender'=>$this->faker->randomElement(['M', 'F']),
            'years'=>$this->faker->numberBetween(18,30),
            'socio_professional'=>$this->faker->domainName,
            'student_university'=>$this->faker->words(15, true),
            'student_speciality'=>$this->faker->words(15, true),
            'e_mail'=>$this->faker->safeEmail(),
            'phone_number'=>$this->faker->phoneNumber(),
            'linkedin'=>$this->faker->sentence(15, true),
            'activites_id'=>$this->faker->numberBetween(1, 20),
            'stages_id'=>$this->faker->numberBetween(1, 20)


        ];
    }
}
