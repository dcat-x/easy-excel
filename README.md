<div align="center">

# Easy Excel

<p>
    <a href="https://github.com/dcat-x/easy-excel/actions"><img src="https://github.com/dcat-x/easy-excel/workflows/Tests/badge.svg" alt="Tests"></a>
    <a href="https://packagist.org/packages/dcat-x/easy-excel"><img src="https://img.shields.io/packagist/v/dcat-x/easy-excel.svg" alt="Latest Version"></a>
    <a href="https://packagist.org/packages/dcat-x/easy-excel"><img src="https://img.shields.io/packagist/dt/dcat-x/easy-excel.svg" alt="Total Downloads"></a>
    <a href="https://packagist.org/packages/dcat-x/easy-excel"><img src="https://img.shields.io/packagist/php-v/dcat-x/easy-excel.svg" alt="PHP Version"></a>
    <a href="https://github.com/dcat-x/easy-excel/blob/main/LICENSE"><img src="https://img.shields.io/packagist/l/dcat-x/easy-excel.svg" alt="License"></a>
</p>

**基于 [OpenSpout](https://github.com/openspout/openspout) 封装的 Excel 读写工具**

</div>

---

Easy Excel 提供简洁优雅的 API 来读写 Excel 文件，无论文件多大都只占用极少内存。

支持格式：`xlsx`、`csv`、`ods`

## 安装

```bash
composer require dcat-x/easy-excel
```

### 环境要求

- PHP >= 8.2
- PHP 扩展：`zip`、`xmlreader`

## 快速开始

### 导出

#### 下载文件

```php
use Dcat\EasyExcel\Excel;

$data = [
    ['id' => 1, 'name' => 'Tom', 'email' => 'tom@example.com'],
    ['id' => 2, 'name' => 'Jerry', 'email' => 'jerry@example.com'],
];

// 设置表头映射
$headings = ['id' => 'ID', 'name' => '姓名', 'email' => '邮箱'];

// 导出为 xlsx
Excel::export($data)->headings($headings)->download('users.xlsx');

// 导出为 csv
Excel::export($data)->headings($headings)->download('users.csv');

// 导出为 ods
Excel::export($data)->headings($headings)->download('users.ods');
```

#### 保存到服务器

```php
use Dcat\EasyExcel\Excel;

// 保存到本地路径
Excel::export($data)->store('/tmp/users.xlsx');

// 使用 Flysystem
use League\Flysystem\Filesystem;
use League\Flysystem\Local\LocalFilesystemAdapter;

$filesystem = new Filesystem(new LocalFilesystemAdapter(__DIR__));

Excel::export($data)->disk($filesystem)->store('users.xlsx');
```

#### 获取文件内容

```php
use Dcat\EasyExcel\Excel;

$xlsxContent = Excel::xlsx($data)->raw();
$csvContent = Excel::csv($data)->raw();
$odsContent = Excel::ods($data)->raw();
```

### 导入

#### 读取所有数据

```php
use Dcat\EasyExcel\Excel;

// 指定表头字段
$headings = ['id', 'name', 'email'];

// 读取所有工作表
$allSheets = Excel::import('/tmp/users.xlsx')->headings($headings)->toArray();
// 返回: ['Sheet1' => [['id' => 1, 'name' => 'Tom', 'email' => 'tom@example.com'], ...]]

// 使用 Flysystem
use League\Flysystem\Filesystem;
use League\Flysystem\Local\LocalFilesystemAdapter;

$filesystem = new Filesystem(new LocalFilesystemAdapter(__DIR__));

$allSheets = Excel::import('users.xlsx')->disk($filesystem)->toArray();
```

#### 读取指定工作表

```php
use Dcat\EasyExcel\Excel;

// 获取第一个工作表
$firstSheet = Excel::import('/tmp/users.xlsx')->first()->toArray();

// 获取最后活动的工作表
$activeSheet = Excel::import('/tmp/users.xlsx')->active()->toArray();

// 按名称或索引获取
$sheet = Excel::import('/tmp/users.xlsx')->sheet('Sheet1')->toArray();
$sheet = Excel::import('/tmp/users.xlsx')->sheet(0)->toArray();
```

#### 遍历工作表

```php
use Dcat\EasyExcel\Excel;
use Dcat\EasyExcel\Contracts\Sheet;
use Dcat\EasyExcel\Support\SheetCollection;

Excel::import('/tmp/users.xlsx')->each(function (Sheet $sheet) {
    $name = $sheet->getName();    // 工作表名称
    $index = $sheet->getIndex();  // 工作表索引 (从 0 开始)

    // 分块处理大数据
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

## License

[MIT License](LICENSE)
