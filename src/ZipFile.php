<?php

namespace Odan\Excel;

use ZipStream\ZipStream;

final class ZipFile implements FileWriterInterface, FileReaderInterface
{
    private ZipStream $zip;

    /**
     * @var resource
     */
    private $stream;

    public function __construct(string $filename = 'php://memory')
    {
        // Create ZIP file, only in-memory
        $this->stream = fopen($filename, 'w+b');

        $this->zip = new ZipStream(
            outputStream: $this->stream,
            sendHttpHeaders: false,
        );
    }

    public function addFile(string $name, string $data): void
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
        $this->zip->finish();

        rewind($this->stream);

        return $this->stream;
    }
}
