<?php

namespace Dcat\EasyExcel;

use Dcat\EasyExcel\Exporters\Exporter;
use Dcat\EasyExcel\Exporters\Sheet;
use Dcat\EasyExcel\Importers\Importer;
use Symfony\Component\HttpFoundation\File\UploadedFile;

class Excel
{
    const XLSX = 'xlsx';

    const CSV = 'csv';

    const ODS = 'ods';

    /**
     * 导入.
     *
     * @param  string|UploadedFile  $filePath
     */
    public static function import($filePath): Contracts\Importer
    {
        return new Importer($filePath);
    }

    /**
     * 导出.
     *
     * @param  array|\Closure|\Generator  $data
     */
    public static function export($data = null): Contracts\Exporter
    {
        return new Exporter($data);
    }

    /**
     * @param  array|\Closure|\Generator  $data
     * @param  null  $sheetName
     * @param  null  $headings
     */
    public static function createSheet($data = null, $sheetName = null, array $headings = []): Contracts\Exporters\Sheet
    {
        return new Sheet($data, $sheetName, $headings);
    }

    /**
     * @param  mixed  ...$params
     * @return Factory
     */
    public static function xlsx(...$params)
    {
        return new Factory(static::XLSX, $params);
    }

    /**
     * @param  mixed  ...$params
     * @return Factory
     */
    public static function csv(...$params)
    {
        return new Factory(static::CSV, $params);
    }

    /**
     * @param  mixed  ...$params
     * @return Factory
     */
    public static function ods(...$params)
    {
        return new Factory(static::ODS, $params);
    }
}
