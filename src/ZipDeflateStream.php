<?php

namespace Odan\Excel;

use RuntimeException;

final class ZipDeflateStream implements ZipStreamWriterInterface, ZipStreamInterface
{
    /**
     * @var resource
     */
    private $stream;

    /** @var array<int, mixed> */
    private array $files = [];

    private bool $isClosed = false;
    private int $deflateLevel;

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
        $this->deflateLevel = $deflateLevel;
    }

    public function write(string $name, string $data): void
    {
        // Compress the data using DEFLATE compression
        // 9 is the highest compression level
        $compressedData = (string)gzdeflate($data, $this->deflateLevel);

        $file = [
            'filename' => $name,
            'checksum' => crc32($data),
            'compressed_size' => strlen($compressedData),
            'uncompressed_size' => strlen($data),
            'filename_length' => strlen($name),
            'relative_offset' => ftell($this->stream),
        ];

        $this->files[] = $file;

        // ZIP file format begins with a local file header
        $localFileHeader = "\x50\x4B\x03\x04"; // Local file header signature
        $localFileHeader .= "\x14\x00"; // Version 20 -> 2.0
        $localFileHeader .= "\x00\x00"; // No Flags
        $localFileHeader .= "\x08\x00"; // Compression method (DEFLATE)
        $localFileHeader .= "\x1c\x7d"; // Last modified time 15:40:56
        $localFileHeader .= "\x4b\x35"; // File last modification date 10/11/2006
        $localFileHeader .= pack('V', $file['checksum']); // CRC32 checksum
        $localFileHeader .= pack('V', $file['compressed_size']); // Compressed size, 4 bytes
        $localFileHeader .= pack('V', $file['uncompressed_size']); // Uncompressed size, 4 bytes
        $localFileHeader .= pack('v', $file['filename_length']); // File name length 2 bytes
        $localFileHeader .= pack('v', 0); // Extra field length
        $localFileHeader .= $name; // File name

        // Write the local file header to the ZIP file
        fwrite($this->stream, $localFileHeader);

        // Write the compressed data to the ZIP file
        fwrite($this->stream, $compressedData);
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
        // ZIP file format ends with the central directory structure and end of central directory record
        $startOfCentralDirectory = ftell($this->stream);
        $centralDirectoryLength = 0;

        // Central directory structure
        foreach ($this->files as $file) {
            $centralDirectory = "\x50\x4B\x01\x02"; // Central directory file header signature
            $centralDirectory .= "\x17\x03"; // Version made by UNIX 2.3
            $centralDirectory .= "\x14\x00"; // Flags 0x14 = 20 -> 2.0
            $centralDirectory .= "\x00\x00"; // No Flags
            $centralDirectory .= "\x08\x00"; // Compression method (DEFLATE)
            $centralDirectory .= "\x1c\x7d\x4b\x35"; // Last modified time
            $centralDirectory .= pack('V', $file['checksum']); // CRC32 checksum
            $centralDirectory .= pack('V', $file['compressed_size']); // Compressed size, 4 bytes
            $centralDirectory .= pack('V', $file['uncompressed_size']); // Uncompressed size, 4 bytes
            $centralDirectory .= pack('v', $file['filename_length']); // File name length, 2 bytes
            $centralDirectory .= pack('v', 0); // Extra field length
            $centralDirectory .= pack('v', 0); // File comment length
            $centralDirectory .= pack('v', 0); // Disk number start
            $centralDirectory .= pack('v', 0); // Internal file attributes. Bit 0 set: ASCII/text file
            $centralDirectory .= "\x00\x00\xa4\x81"; // External file attributes (regular file)
            // Relative offset of local header.
            // This is the offset of where to find the corresponding
            // local file header from the start of the first disk.
            $centralDirectory .= pack('V', $file['relative_offset']);
            $centralDirectory .= $file['filename']; // File name

            // Write the central directory structure to the ZIP file
            fwrite($this->stream, $centralDirectory);

            $centralDirectoryLength = $centralDirectoryLength + strlen($centralDirectory);
        }

        $numberOfEntries = count($this->files);

        // End of central directory record
        $endOfCentralDirectory = "\x50\x4B\x05\x06"; // End of central directory record signature
        $endOfCentralDirectory .= "\x00\x00"; // Number of this disk
        $endOfCentralDirectory .= "\x00\x00"; // Number of the disk with the start of the central directory
        // Total number of entries in the central directory on this disk
        $endOfCentralDirectory .= pack('v', $numberOfEntries);
        // Total number of entries in the central directory
        $endOfCentralDirectory .= pack('v', $numberOfEntries);
        // Size of the central directory
        $endOfCentralDirectory .= pack('V', $centralDirectoryLength);
        // Offset of start of central directory with respect to the starting disk number
        $endOfCentralDirectory .= pack('V', $startOfCentralDirectory);
        $endOfCentralDirectory .= "\x00\x00"; // ZIP file comment length

        // Write the end of central directory record to the ZIP file
        fwrite($this->stream, $endOfCentralDirectory);

        rewind($this->stream);

        $this->isClosed = true;
    }
}
