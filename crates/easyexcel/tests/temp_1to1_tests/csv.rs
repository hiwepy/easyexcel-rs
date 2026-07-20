//! 1:1 method matrix for Java `com.alibaba.easyexcel.test.temp.csv.*`

use super::helpers;

/// Java `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#write`
#[test]
fn csv_csv_read_test_write() {
    helpers::assert_csv_write_read("1to1_csv.csv");
}

/// Java `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#read1`
#[test]
fn csv_csv_read_test_read_1() {
    helpers::assert_csv_fixture_read();
}

/// Java `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#csvWrite`
#[test]
fn csv_csv_read_test_csv_write() {
    helpers::assert_csv_write_read("1to1_csv.csv");
}

/// Java `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#writev2`
#[test]
fn csv_csv_read_test_writev_2() {
    helpers::assert_csv_write_read("1to1_csv.csv");
}

/// Java `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#writeFile`
#[test]
fn csv_csv_read_test_write_file() {
    helpers::assert_csv_file_magic();
}

/// Java `com.alibaba.easyexcel.test.temp.csv.CsvReadTest#read`
#[test]
fn csv_csv_read_test_read() {
    helpers::assert_csv_fixture_read();
}
