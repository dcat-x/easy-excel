<?php

namespace Dcat\EasyExcel\Contracts;

use Dcat\EasyExcel\Support\SheetCollection;
use OpenSpout\Reader\SheetInterface;

interface Sheet
{
    public function valid(): bool;

    /**
     * sheet索引.
     *
     * @return int
     */
    public function getIndex();

    /**
     * @return string
     */
    public function getName();

    /**
     * 获取原始标题.
     *
     * @return array
     */
    public function getOriginalHeadings();

    /**
     * @return bool
     */
    public function isActive();

    /**
     * @return bool
     */
    public function isVisible();

    /**
     * @return SheetInterface
     */
    public function getSheet();

    /**
     * @return $this
     */
    public function filter(callable $callback);

    /**
     * 逐行读取.
     *
     * e.g:
     *
     * $this->each(function (array $row, $k, $headers) {
     *      ...
     * });
     *
     * @param  callable|null  $callback
     * @return $this
     */
    public function each(callable $callback);

    /**
     * 分块读取.
     *
     * e.g:
     *
     * $this->chunk(100, function (SheetCollection $collection) {
     *      ...
     * });
     *
     * @return \Dcat\EasyExcel\Importers\Sheet
     */
    public function chunk(int $size, callable $callback);

    public function toArray(): array;

    public function collect(): SheetCollection;
}
