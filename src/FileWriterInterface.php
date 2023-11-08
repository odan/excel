<?php

namespace Odan\Excel;

interface FileWriterInterface
{
    public function write(string $name, string $data): void;
}
