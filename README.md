# parsingExcel
Library ini berguna untuk parsing excel menggunakan PHPExcel.
cara penggunaan :
# Untuk Parsing All Sheet
```php
parsingExcel::parsingExcel($dir_file);
```
# Untuk Memilih Sheet yang di inginkan
```php
$status = true;
$numberSheet = 0;
parsingExcel::parsingExcel($dir_file,$status,$numberSheet);
```
