## TEMPLATE KAPAL RUTE
1. ini digunakan untuk mengatur kapal dan urutan kapal setiap rute diexcel untuk produksi
2. ini akan berguna diexport laporan asdp, ketika ada kapal asdp dirute tersebut, maka akan menampilkan rute
3. jika kamu mengexport db nya dari database, terus di import kembali, kadang ada error karena order by, yang perlu kamu lakukan adalah mengganti

``` sql
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `v_laporan_produksi_harian`  AS SELECT `p`.`produksi_id` AS `produksi_id`, `p`.`tanggal_produksi` AS `tanggal_produksi`, `p`.`nama_perusahaan` AS `nama_perusahaan`, `p`.`nama_kapal` AS `nama_kapal`, `p`.`nama_rute` AS `nama_rute`, `p`.`shift` AS `shift`, `p`.`regu` AS `regu`, `p`.`total_penumpang` AS `total_penumpang`, `p`.`total_pendapatan_penumpang` AS `total_pendapatan_penumpang`, `p`.`total_kendaraan` AS `total_kendaraan`, `p`.`total_pendapatan_kendaraan` AS `total_pendapatan_kendaraan`, `p`.`total_pendapatan` AS `total_pendapatan`, `u`.`nama_lengkap` AS `created_by_name`, `p`.`created_at` AS `created_at`, `p`.`updated_at` AS `updated_at` FROM (`produksi` `p` left join `users` `u` on((`p`.`created_by` = `u`.`user_id`))) ORDER BY `p`.`tanggal_produksi` DESC, `p`.`created_at` AS `DESCdesc` ASC  ;
```
ganti dengan 
``` sql
CREATE ALGORITHM=UNDEFINED DEFINER=`root`@`localhost` SQL SECURITY DEFINER VIEW `v_laporan_produksi_harian`  AS SELECT `p`.`produksi_id` AS `produksi_id`, `p`.`tanggal_produksi` AS `tanggal_produksi`, `p`.`nama_perusahaan` AS `nama_perusahaan`, `p`.`nama_kapal` AS `nama_kapal`, `p`.`nama_rute` AS `nama_rute`, `p`.`shift` AS `shift`, `p`.`regu` AS `regu`, `p`.`total_penumpang` AS `total_penumpang`, `p`.`total_pendapatan_penumpang` AS `total_pendapatan_penumpang`, `p`.`total_kendaraan` AS `total_kendaraan`, `p`.`total_pendapatan_kendaraan` AS `total_pendapatan_kendaraan`, `p`.`total_pendapatan` AS `total_pendapatan`, `u`.`nama_lengkap` AS `created_by_name`, `p`.`created_at` AS `created_at`, `p`.`updated_at` AS `updated_at` FROM (`produksi` `p` left join `users` `u` on((`p`.`created_by` = `u`.`user_id`))) ORDER BY `p`.`tanggal_produksi` DESC, `p`.`created_at` DESC  ;
```
