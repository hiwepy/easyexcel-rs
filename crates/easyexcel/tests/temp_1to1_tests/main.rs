//! Temp package method-level 1:1 naming matrix.
//!
//! Every Java `@Test` under `com.alibaba.easyexcel.test.temp` maps to
//! `#[test] fn <pkg_snake>_<class_snake>_<method_snake>`.
//! Portable cases assert via EasyExcel APIs; non-portable cases are
//! `#[ignore]` placeholders so the method name count stays 1:1.
//!
//! Matrix: `docs/temp-1to1-matrix.md`.

mod cache;
mod csv;
mod dataformat;
mod fill;
mod helpers;
mod issue1662;
mod issue1663;
mod issue2443;
mod large;
mod poi;
mod read;
mod root;
mod simple;
mod write;
