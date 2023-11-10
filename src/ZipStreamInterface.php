<?php

namespace Odan\Excel;

interface ZipStreamInterface
{
    /**
     * @return resource
     */
    public function getStream(): mixed;
}
