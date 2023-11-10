<?php

namespace Odan\Excel;

use RuntimeException;
use ZipStream\ZipStream;

final class Zip64Stream implements ZipStreamWriterInterface, ZipStreamInterface
{
    private ZipStream $zip;

    /**
     * @var resource
     */
    private $stream;

    private bool $isClosed = false;

    public function __construct(string $filename = 'php://memory', int $deflateLevel = 6)
    {
        // w+b: If the file does not exist, it will be created.
        // If it already exists, its content will be truncated (cleared)
        // when you write data to it.
        $stream = fopen($filename, 'w+b');
        if ($stream === false) {
            throw new RuntimeException('File could not be opened.');
        }

        $this->stream = $stream;

        // defaultEnableZeroHeader must be set to false to generate Excel compatible ZIP files
        $this->zip = new ZipStream(
            defaultEnableZeroHeader: false,
            defaultDeflateLevel: $deflateLevel,
            outputStream: $this->stream,
            sendHttpHeaders: false,
        );
    }

    public function write(string $name, string $data): void
    {
        $this->zip->addFile(
            fileName: $name,
            data: $data,
        );
    }

    /**
     * @return resource
     */
    public function getStream(): mixed
    {
        if (!$this->isClosed) {
            $this->close();
        }

        rewind($this->stream);

        return $this->stream;
    }

    public function close(): void
    {
        $this->zip->finish();
        $this->isClosed = true;
    }
}
