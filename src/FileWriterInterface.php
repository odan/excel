<?php

namespace Odan\Excel;

interface FileWriterInterface
{
    public function addFile(string $name, string $data): void;
}
