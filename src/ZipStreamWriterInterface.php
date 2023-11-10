<?php

namespace Odan\Excel;

interface ZipStreamWriterInterface
{
    public function write(string $name, string $data): void;

    public function close(): void;
}
