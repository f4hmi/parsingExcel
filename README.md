# parsingExcel
Library ini berguna untuk parsing excel menggunakan PHPExcel.
cara penggunaan :
# Untuk Parsing All Sheet
```php
$status = true;
$result = parsingExcel::parsingExcel($dir_file);
```
# Untuk Memilih Sheet yang di inginkan
```php
$status = false;
$numberSheet = 0;
$result = parsingExcel::parsingExcel($dir_file,$status,$numberSheet);
```
# Result Allsheet
```array
Array
(
    [sheet_0] => Array
        (
            [0] => Array
                (
                    [0] => designator
                    [1] => uraian
                    [2] => type
                    [3] => satuan
                    [4] => no_urut
                )

            [1] => Array
                (
                    [0] => DC-OF-SM-12D
                    [1] => Pengadaan dan pemasangan Kabel Duct Fiber Optik Single Mode 12 core G 652 D
                    [2] => barang & jasa
                    [3] => meter
                    [4] => 
                )

            [2] => Array
                (
                    [0] => DC-OF-SM-24D
                    [1] => Pengadaan dan pemasangan Kabel Duct Fiber Optik Single Mode 24 core G 652 D 
                    [2] => barang & jasa
                    [3] => meter
                    [4] => 
                )
         }
}
```
# Result Sheet yang dipilih
```
Array
(
  [0] => Array
      (
          [0] => designator
          [1] => uraian
          [2] => type
          [3] => satuan
          [4] => no_urut
      )

  [1] => Array
      (
          [0] => DC-OF-SM-12D
          [1] => Pengadaan dan pemasangan Kabel Duct Fiber Optik Single Mode 12 core G 652 D
          [2] => barang & jasa
          [3] => meter
          [4] => 
      )

  [2] => Array
      (
          [0] => DC-OF-SM-24D
          [1] => Pengadaan dan pemasangan Kabel Duct Fiber Optik Single Mode 24 core G 652 D 
          [2] => barang & jasa
          [3] => meter
          [4] => 
      )
}
```
