<?php

namespace Dcat\EasyExcel\Contracts\Exporters;

use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;

interface Sheet
{
    /**
     * @return $this
     */
    public function data($data);

    /**
     * @return $this
     */
    public function chunk(callable $callback);

    /**
     * @return array|\Generator
     */
    public function getData();

    /**
     * 传false则禁用标题.
     *
     * @param  array|false  $headings
     * @return $this
     */
    public function headings($headings);

    /**
     * @return array|false
     */
    public function getHeadings();

    /**
     * @param  Style  $style
     * @return $this
     */
    public function headingStyle($style);

    /**
     * @return Style
     */
    public function getHeadingStyle();

    /**
     * @return $this
     */
    public function name(?string $name);

    /**
     * @return string
     */
    public function getName();

    /**
     * @return $this
     */
    public function row(\Closure $callback);

    /**
     * @return array|Row
     */
    public function formatRow(array $row);
}
