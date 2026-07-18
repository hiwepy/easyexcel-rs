use std::cell::RefCell;
use std::fs::File;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::PathBuf;
use std::thread;

use easyexcel_core::{ExcelError, Result};
use tempfile::NamedTempFile;

/// Shared-string cache selection equivalent to Java `EasyExcel`'s built-in
/// `ReadCacheSelector` behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ReadCacheMode {
    /// Keep small `sharedStrings.xml` parts in memory and spill larger parts to disk.
    #[default]
    Auto,
    /// Keep every shared string in memory, equivalent to Java's `MapCache`.
    Memory,
    /// Store UTF-8 shared-string values in a temporary file.
    Disk,
}

pub(crate) const DEFAULT_MAX_MEMORY_SHARED_STRINGS_BYTES: u64 = 5_000_000;

/// Write-phase cache: sequential `put` calls.
pub(crate) trait SharedStringCacheWriter {
    fn put(&mut self, value: String) -> Result<()>;
    fn finish(self: Box<Self>) -> Result<Box<dyn SharedStringCacheReader>>;
}

/// Read-phase cache: concurrent `get` calls via `&self` (no `&mut`).
pub(crate) trait SharedStringCacheReader: Send + Sync {
    fn get(&self, index: usize) -> Result<String>;
    #[cfg(test)]
    fn len(&self) -> usize;
}

/// Unified trait for backward compatibility in tests.
pub(crate) trait SharedStringCache: SharedStringCacheWriter + SharedStringCacheReader {
    fn put_and_finish(mut self: Box<Self>) -> Result<Box<dyn SharedStringCacheReader>>
    where
        Self: Sized,
    {
        self.finish()
    }
}

pub(crate) fn memory_cache() -> Box<dyn SharedStringCacheReader> {
    Box::new(MemorySharedStringReader::default())
}

pub(crate) fn create_cache(
    mode: ReadCacheMode,
    xml_size: u64,
) -> Result<Box<dyn SharedStringCache>> {
    match mode {
        ReadCacheMode::Auto if xml_size < DEFAULT_MAX_MEMORY_SHARED_STRINGS_BYTES => {
            Ok(Box::new(MemorySharedStringCache::default()))
        }
        ReadCacheMode::Auto | ReadCacheMode::Disk => {
            box_disk_cache(ConcurrentDiskCache::new())
        }
        ReadCacheMode::Memory => Ok(Box::new(MemorySharedStringCache::default())),
    }
}

// ---------------------------------------------------------------------------
// MemorySharedStringCache — all data in RAM, no thread_local needed
// ---------------------------------------------------------------------------

#[derive(Default)]
struct MemorySharedStringCache {
    values: Vec<String>,
}

impl SharedStringCacheWriter for MemorySharedStringCache {
    fn put(&mut self, value: String) -> Result<()> {
        self.values.push(value);
        Ok(())
    }

    fn finish(self: Box<Self>) -> Result<Box<dyn SharedStringCacheReader>> {
        Ok(Box::new(MemorySharedStringReader { values: self.values }))
    }
}

impl SharedStringCacheReader for MemorySharedStringCache {
    fn get(&self, index: usize) -> Result<String> {
        self.values.get(index).cloned().ok_or_else(|| {
            ExcelError::Format(format!("shared string index is out of bounds: {index}"))
        })
    }

    #[cfg(test)]
    fn len(&self) -> usize {
        self.values.len()
    }
}

impl SharedStringCache for MemorySharedStringCache {}

/// Read-only view after `put_finished()`.
#[derive(Default)]
struct MemorySharedStringReader {
    values: Vec<String>,
}

impl SharedStringCacheReader for MemorySharedStringReader {
    fn get(&self, index: usize) -> Result<String> {
        self.values.get(index).cloned().ok_or_else(|| {
            ExcelError::Format(format!("shared string index is out of bounds: {index}"))
        })
    }

    #[cfg(test)]
    fn len(&self) -> usize {
        self.values.len()
    }
}

// ---------------------------------------------------------------------------
// ConcurrentDiskCache — thread_local! per-thread readers
// ---------------------------------------------------------------------------

/// Disk-backed shared-string cache with concurrent read support.
///
/// Write phase: sequential `put` calls write to the temp file.
/// Read phase: each thread opens its own `File` handle via `thread_local!`,
/// so `get(&self)` requires no `&mut` and supports full parallelism.
struct ConcurrentDiskCache {
    _temporary_file: Option<NamedTempFile>,
    writer: Option<Box<dyn WriteSeek>>,
    /// File path for per-thread readers.
    path: PathBuf,
    /// Shared, read-only index (offset, length) for each string.
    entries: Vec<(u64, usize)>,
}

trait WriteSeek: Write + Seek + Send + Sync {}
impl<T: Write + Seek + Send + Sync> WriteSeek for T {}

thread_local! {
    static TLS_READER: RefCell<Option<File>> = const { RefCell::new(None) };
}

impl ConcurrentDiskCache {
    fn new() -> Result<Self> {
        Self::from_temporary_file(NamedTempFile::new())
    }

    fn from_temporary_file(temporary_file: std::io::Result<NamedTempFile>) -> Result<Self> {
        match temporary_file {
            Ok(temporary_file) => {
                let path = temporary_file.path().to_path_buf();
                let writer = temporary_file.reopen();
                Self::from_parts(temporary_file, path, writer)
            }
            Err(error) => Err(error.into()),
        }
    }

    fn from_parts(
        temporary_file: NamedTempFile,
        path: PathBuf,
        writer: std::io::Result<File>,
    ) -> Result<Self> {
        match writer {
            Ok(writer) => Ok(Self {
                _temporary_file: Some(temporary_file),
                writer: Some(Box::new(writer)),
                path,
                entries: Vec::new(),
            }),
            Err(error) => Err(error.into()),
        }
    }
}

impl SharedStringCacheWriter for ConcurrentDiskCache {
    fn put(&mut self, value: String) -> Result<()> {
        let writer = self.writer.as_mut().ok_or_else(|| {
            ExcelError::Format("cache writer already finished".to_owned())
        })?;
        let offset = writer.seek(SeekFrom::End(0))?;
        let bytes = value.as_bytes();
        writer.write_all(bytes)?;
        self.entries.push((offset, bytes.len()));
        Ok(())
    }

    fn finish(mut self: Box<Self>) -> Result<Box<dyn SharedStringCacheReader>> {
        let path = self.path.clone();
        let entries = self.entries.clone();
        // Take the writer out so the read-only view is Send + Sync
        let _ = self.writer.take();
        Ok(Box::new(ConcurrentDiskReader {
            _temporary_file: self._temporary_file,
            path,
            entries,
        }))
    }
}

impl SharedStringCacheReader for ConcurrentDiskCache {
    fn get(&self, index: usize) -> Result<String> {
        let (offset, length) = self.entries.get(index).copied().ok_or_else(|| {
            ExcelError::Format(format!("shared string index is out of bounds: {index}"))
        })?;
        TLS_READER.with(|cell| {
            let mut guard = cell.borrow_mut();
            let file = guard.get_or_insert_with(|| File::open(&self.path).unwrap());
            file.seek(SeekFrom::Start(offset))?;
            let mut bytes = vec![0u8; length];
            file.read_exact(&mut bytes)?;
            String::from_utf8(bytes).map_err(|error| ExcelError::Format(error.to_string()))
        })
    }

    #[cfg(test)]
    fn len(&self) -> usize {
        self.entries.len()
    }
}

impl SharedStringCache for ConcurrentDiskCache {}

/// Read-only view after `put_finished()` — uses thread-local File handles.
struct ConcurrentDiskReader {
    _temporary_file: Option<NamedTempFile>,
    path: PathBuf,
    entries: Vec<(u64, usize)>,
}

impl SharedStringCacheReader for ConcurrentDiskReader {
    fn get(&self, index: usize) -> Result<String> {
        let (offset, length) = self.entries.get(index).copied().ok_or_else(|| {
            ExcelError::Format(format!("shared string index is out of bounds: {index}"))
        })?;
        TLS_READER.with(|cell| {
            let mut guard = cell.borrow_mut();
            let file = guard.get_or_insert_with(|| File::open(&self.path).unwrap());
            file.seek(SeekFrom::Start(offset))?;
            let mut bytes = vec![0u8; length];
            file.read_exact(&mut bytes)?;
            String::from_utf8(bytes).map_err(|error| ExcelError::Format(error.to_string()))
        })
    }

    #[cfg(test)]
    fn len(&self) -> usize {
        self.entries.len()
    }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn box_disk_cache(cache: Result<ConcurrentDiskCache>) -> Result<Box<dyn SharedStringCache>> {
    match cache {
        Ok(cache) => Ok(Box::new(cache)),
        Err(error) => Err(error),
    }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    fn io_error() -> std::io::Error {
        std::io::Error::other("injected cache I/O failure")
    }

    struct FaultyIo {
        fail_seek: bool,
        fail_read: bool,
        fail_write: bool,
    }

    impl Seek for FaultyIo {
        fn seek(&mut self, _position: SeekFrom) -> std::io::Result<u64> {
            if self.fail_seek {
                Err(io_error())
            } else {
                Ok(0)
            }
        }
    }

    impl Read for FaultyIo {
        fn read(&mut self, _buffer: &mut [u8]) -> std::io::Result<usize> {
            if self.fail_read {
                Err(io_error())
            } else {
                Ok(0)
            }
        }
    }

    impl Write for FaultyIo {
        fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
            if self.fail_write {
                Err(io_error())
            } else {
                Ok(buffer.len())
            }
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn memory_cache_preserves_utf8_values() {
        let mut cache = MemorySharedStringCache::default();
        assert_eq!(cache.len(), 0);
        cache.put("alpha".to_owned()).expect("first value");
        cache.put("中文😀".to_owned()).expect("Unicode value");
        assert_eq!(cache.len(), 2);
        assert_eq!(cache.get(0).expect("first value"), "alpha");
        assert_eq!(cache.get(1).expect("Unicode value"), "中文😀");
        assert!(cache.get(2).is_err());
    }

    #[test]
    fn disk_cache_preserves_utf8_values() {
        let mut cache = ConcurrentDiskCache::new().expect("disk cache");
        assert_eq!(cache.len(), 0);
        cache.put("alpha".to_owned()).expect("first value");
        cache.put("中文😀".to_owned()).expect("Unicode value");
        assert_eq!(cache.len(), 2);
        assert_eq!(cache.get(0).expect("first value"), "alpha");
        assert_eq!(cache.get(1).expect("Unicode value"), "中文😀");
        assert!(cache.get(2).is_err());
    }

    #[test]
    fn disk_cache_concurrent_reads() {
        let mut cache = ConcurrentDiskCache::new().expect("disk cache");
        for i in 0..100 {
            cache.put(format!("value-{i}")).expect("put");
        }

        // Spawn multiple threads reading concurrently
        let mut handles = vec![];
        for thread_id in 0..4 {
            let reader_clone = ConcurrentDiskReader {
                _temporary_file: None,
                path: cache.path.clone(),
                entries: cache.entries.clone(),
            };
            handles.push(thread::spawn(move || {
                for i in 0..100 {
                    let value = reader_clone.get(i).expect("get");
                    assert_eq!(value, format!("value-{i}"));
                }
                thread_id
            }));
        }
        for h in handles {
            h.join().expect("thread joined");
        }
    }

    #[test]
    fn automatic_cache_selection_uses_java_five_megabyte_boundary() {
        let mut memory = create_cache(ReadCacheMode::Auto, 4_999_999).expect("memory cache");
        memory.put("memory".to_owned()).expect("memory value");
        assert_eq!(memory.get(0).expect("memory value"), "memory");

        let mut disk = create_cache(ReadCacheMode::Auto, 5_000_000).expect("disk cache");
        disk.put("disk".to_owned()).expect("disk value");
        assert_eq!(disk.get(0).expect("disk value"), "disk");
        assert_eq!(ReadCacheMode::default(), ReadCacheMode::Auto);
    }

    #[test]
    fn disk_cache_propagates_creation_failures() {
        assert!(ConcurrentDiskCache::from_temporary_file(Err(io_error())).is_err());
        let temporary_file = NamedTempFile::new().expect("temporary file");
        let path = temporary_file.path().to_path_buf();
        assert!(ConcurrentDiskCache::from_parts(temporary_file, path, Err(io_error())).is_err());
        assert!(box_disk_cache(Err(ExcelError::Format("injected".to_owned()))).is_err());
    }

    #[test]
    fn disk_cache_propagates_write_failures() {
        let temporary_file = NamedTempFile::new().expect("temporary file");
        let path = temporary_file.path().to_path_buf();
        let mut cache = ConcurrentDiskCache {
            _temporary_file: Some(temporary_file),
            writer: Some(Box::new(FaultyIo { fail_seek: true, fail_read: false, fail_write: false })),
            path,
            entries: Vec::new(),
        };
        assert!(cache.put("value".to_owned()).is_err());

        let temporary_file = NamedTempFile::new().expect("temporary file");
        let path = temporary_file.path().to_path_buf();
        let mut cache2 = ConcurrentDiskCache {
            _temporary_file: Some(temporary_file),
            writer: Some(Box::new(FaultyIo { fail_seek: false, fail_read: false, fail_write: true })),
            path,
            entries: Vec::new(),
        };
        assert!(cache2.put("value".to_owned()).is_err());
    }

    #[test]
    fn disk_cache_propagates_read_failures() {
        let temporary_file = NamedTempFile::new().expect("temporary file");
        let path = temporary_file.path().to_path_buf();
        let mut cache = ConcurrentDiskCache::new().expect("disk cache");
        cache.put("x".to_owned()).expect("cache value");

        // Corrupt the file to trigger a read error
        let writer = cache.writer.as_mut().unwrap();
        writer.seek(SeekFrom::Start(0)).expect("seek");
        writer.write_all(&[0xff]).expect("corrupt value");
        assert!(cache.get(0).is_err());
    }
}
