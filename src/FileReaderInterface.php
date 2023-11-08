<?php

namespace Odan\Excel;

interface FileReaderInterface
{
    /**
     * @return resource
     */
    public function readStream(): mixed;
}
