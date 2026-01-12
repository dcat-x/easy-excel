<?php

namespace Tests;

use PHPUnit\Framework\TestCase as BaseTestCase;

abstract class TestCase extends BaseTestCase
{
    /**
     * @return array
     */
    protected function convertDatetimeObjectToString(array $values)
    {
        foreach ($values as &$value) {
            if ($value instanceof \DateTimeInterface) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif (is_int($value) || is_float($value)) {
                $value = (string) $value;
            }
        }

        return $values;
    }
}
