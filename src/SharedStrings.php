<?php

namespace Odan\Excel;

final class SharedStrings
{
    /** @var array<string, int> */
    private array $sharedStrings = [];

    public function add(string $string): int
    {
        $index = $this->sharedStrings[$string] ?? null;
        if ($index !== null) {
            return $index;
        }

        $newIndex = count($this->sharedStrings);
        $this->sharedStrings[$string] = $newIndex;

        return $newIndex;
    }

    /** @return array<string, int> */
    public function all(): array
    {
        return $this->sharedStrings;
    }
}
