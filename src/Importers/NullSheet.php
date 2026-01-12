<?php

namespace Dcat\EasyExcel\Importers;

use Dcat\EasyExcel\Contracts;
use Dcat\EasyExcel\Support\SheetCollection;
use OpenSpout\Reader\SheetInterface;

class NullSheet implements Contracts\Sheet
{
    public function valid(): bool
    {
        return false;
    }

    /**
     * @return int
     */
    public function getIndex() {}

    /**
     * @return string
     */
    public function getName() {}

    /**
     * @return array
     */
    public function getOriginalHeadings()
    {
        return [];
    }

    /**
     * @return bool
     */
    public function isActive()
    {
        return false;
    }

    /**
     * @return bool
     */
    public function isVisible()
    {
        return false;
    }

    /**
     * @return SheetInterface
     */
    public function getSheet() {}

    /**
     * @return \Dcat\EasyExcel\Contracts\Sheet
     */
    public function filter(callable $callback)
    {
        return $this;
    }

    /**
     * 逐行读取.
     *
     * @param  callable|null  $callback
     * @return $this
     */
    public function each(callable $callback)
    {
        return $this;
    }

    public function chunk(int $size, callable $callback)
    {
        return $this;
    }

    public function toArray(): array
    {
        return [];
    }

    public function collect(): SheetCollection
    {
        return new SheetCollection($this->toArray());
    }
}
