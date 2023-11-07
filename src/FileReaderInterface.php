<?php

namespace Odan\Excel;

interface FileReaderInterface
{
    /**
     * @return resource
     */
    public function getStream(): mixed;
}
