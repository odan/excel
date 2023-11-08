<?php

namespace Odan\Excel;

use UnexpectedValueException;
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
        $stream = fopen($filename, 'w+b');
        if ($stream === false) {
            throw new UnexpectedValueException('File could not be opened.');
        }

        $this->stream = $stream;

        // defaultEnableZeroHeader must be set to false for Excel compatible ZIP files
        $this->zip = new ZipStream(
            defaultEnableZeroHeader: false,
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
