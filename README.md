<div align="center">

# Easy Excel

[![Tests](https://github.com/dcat-x/easy-excel/workflows/Tests/badge.svg)](https://github.com/dcat-x/easy-excel/actions)
[![Latest Version on Packagist](https://img.shields.io/packagist/v/dcat-x/easy-excel.svg)](https://packagist.org/packages/dcat-x/easy-excel)
[![Total Downloads](https://img.shields.io/packagist/dt/dcat-x/easy-excel.svg)](https://packagist.org/packages/dcat-x/easy-excel)
[![PHP Version](https://img.shields.io/packagist/php-v/dcat-x/easy-excel.svg)](https://packagist.org/packages/dcat-x/easy-excel)
[![License](https://img.shields.io/packagist/l/dcat-x/easy-excel.svg)](https://github.com/dcat-x/easy-excel/blob/main/LICENSE)

基于 [OpenSpout](https://github.com/openspout/openspout) 封装的 Excel 读写工具，内存占用极低。

[安装](#安装) |
[使用](#使用) |
[更新日志](CHANGELOG.md) |
[贡献指南](CONTRIBUTING.md)

</div>

---

## 特性

- 读写 Excel 文件，内存占用极低
- 支持 `xlsx`、`csv`、`ods` 格式
- 分块处理，轻松应对大文件
- 集成 Flysystem，灵活存储
- 支持自定义表头和多工作表
- 简洁直观的 API

## 安装

```bash
composer require dcat-x/easy-excel
```

### 环境要求

- PHP >= 8.2
- PHP 扩展：`zip`、`xmlreader`

## 使用

### 导出

#### 下载文件

```php
use Dcat\EasyExcel\Excel;

$data = [
    ['id' => 1, 'name' => 'Tom', 'email' => 'tom@example.com'],
    ['id' => 2, 'name' => 'Jerry', 'email' => 'jerry@example.com'],
];

// 自定义表头
$headings = ['id' => 'ID', 'name' => '姓名', 'email' => '邮箱'];

Excel::export($data)->headings($headings)->download('users.xlsx');
Excel::export($data)->headings($headings)->download('users.csv');
Excel::export($data)->headings($headings)->download('users.ods');
```

#### 保存到服务器

```php
use Dcat\EasyExcel\Excel;
use League\Flysystem\Filesystem;
use League\Flysystem\Local\LocalFilesystemAdapter;

// 保存到本地路径
Excel::export($data)->store('/tmp/users.xlsx');

// 使用 Flysystem
$filesystem = new Filesystem(new LocalFilesystemAdapter(__DIR__));
Excel::export($data)->disk($filesystem)->store('users.xlsx');
```

#### 获取文件内容

```php
use Dcat\EasyExcel\Excel;

$xlsx = Excel::xlsx($data)->raw();
$csv = Excel::csv($data)->raw();
$ods = Excel::ods($data)->raw();
```

### 导入

#### 读取所有数据

```php
use Dcat\EasyExcel\Excel;

$headings = ['id', 'name', 'email'];

$allSheets = Excel::import('/tmp/users.xlsx')->headings($headings)->toArray();
// 返回: ['Sheet1' => [['id' => 1, 'name' => 'Tom', ...], ...]]
```

#### 读取指定工作表

```php
use Dcat\EasyExcel\Excel;

// 第一个工作表
$data = Excel::import('/tmp/users.xlsx')->first()->toArray();

// 最后活动的工作表
$data = Excel::import('/tmp/users.xlsx')->active()->toArray();

// 按名称或索引获取
$data = Excel::import('/tmp/users.xlsx')->sheet('Sheet1')->toArray();
$data = Excel::import('/tmp/users.xlsx')->sheet(0)->toArray();
```

#### 遍历工作表

```php
use Dcat\EasyExcel\Excel;
use Dcat\EasyExcel\Contracts\Sheet;
use Dcat\EasyExcel\Support\SheetCollection;

Excel::import('/tmp/users.xlsx')->each(function (Sheet $sheet) {
    $name = $sheet->getName();
    $index = $sheet->getIndex();

    $sheet->chunk(1000, function (SheetCollection $collection) {
        foreach ($collection as $row) {
            // 处理每一行
        }
    });
});
```

#### 分块处理

```php
use Dcat\EasyExcel\Excel;
use Dcat\EasyExcel\Support\SheetCollection;

// 每次处理 1000 行，适合大文件
Excel::import('/tmp/users.xlsx')
    ->first()
    ->chunk(1000, function (SheetCollection $collection) {
        $data = $collection->toArray();
        // 批量处理数据...
    });
```

## 测试

```bash
composer test
```

## 更新日志

详见 [CHANGELOG](CHANGELOG.md)。

## 贡献

详见 [CONTRIBUTING](CONTRIBUTING.md)。

## 许可证

MIT 许可证，详见 [LICENSE](LICENSE)。
