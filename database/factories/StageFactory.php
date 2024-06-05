<?php

namespace Database\Factories;

use Illuminate\Database\Eloquent\Factories\Factory;

/**
 * @extends \Illuminate\Database\Eloquent\Factories\Factory<\App\Models\Stage>
 */
class StageFactory extends Factory
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
            'projets_solutions' => $this->faker->sentence(),
            'place'=>$this->faker->words(12, true),
            'periode'=>$this->faker->numberBetween(1, 100),
            'date_start'=>$this->faker->date(),
            'date_end'=>$this->faker->date(),
        ];
    }
}
