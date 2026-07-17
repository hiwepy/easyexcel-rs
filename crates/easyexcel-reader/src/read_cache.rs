use std::fs::File;
use std::io::{Read, Seek, SeekFrom, Write};

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

pub(crate) fn memory_cache() -> Box<dyn SharedStringCache> {
    Box::new(MemorySharedStringCache::default())
}

pub(crate) trait SharedStringCache {
    fn put(&mut self, value: String) -> Result<()>;
    fn get(&mut self, index: usize) -> Result<String>;
    #[cfg(test)]
    fn len(&self) -> usize;
}

pub(crate) fn create_cache(
    mode: ReadCacheMode,
    xml_size: u64,
) -> Result<Box<dyn SharedStringCache>> {
    match mode {
        ReadCacheMode::Auto if xml_size < DEFAULT_MAX_MEMORY_SHARED_STRINGS_BYTES => {
            Ok(memory_cache())
        }
        ReadCacheMode::Auto | ReadCacheMode::Disk => box_disk_cache(DiskSharedStringCache::new()),
        ReadCacheMode::Memory => Ok(memory_cache()),
    }
}

#[derive(Default)]
struct MemorySharedStringCache {
    values: Vec<String>,
}

impl SharedStringCache for MemorySharedStringCache {
    fn put(&mut self, value: String) -> Result<()> {
        self.values.push(value);
        Ok(())
    }

    fn get(&mut self, index: usize) -> Result<String> {
        self.values.get(index).cloned().ok_or_else(|| {
            ExcelError::Format(format!("shared string index is out of bounds: {index}"))
        })
    }

    #[cfg(test)]
    fn len(&self) -> usize {
        self.values.len()
    }
}

trait WriteSeek: Write + Seek {}

impl<T: Write + Seek> WriteSeek for T {}

trait ReadSeek: Read + Seek {}

impl<T: Read + Seek> ReadSeek for T {}

struct DiskSharedStringCache {
    _temporary_file: Option<NamedTempFile>,
    writer: Box<dyn WriteSeek>,
    reader: Box<dyn ReadSeek>,
    entries: Vec<(u64, usize)>,
}

impl DiskSharedStringCache {
    fn new() -> Result<Self> {
        Self::from_temporary_file(NamedTempFile::new())
    }

    fn from_temporary_file(temporary_file: std::io::Result<NamedTempFile>) -> Result<Self> {
        match temporary_file {
            Ok(temporary_file) => {
                let writer = temporary_file.reopen();
                let reader = temporary_file.reopen();
                Self::from_handles(temporary_file, writer, reader)
            }
            Err(error) => Err(error.into()),
        }
    }

    fn from_handles(
        temporary_file: NamedTempFile,
        writer: std::io::Result<File>,
        reader: std::io::Result<File>,
    ) -> Result<Self> {
        match (writer, reader) {
            (Ok(writer), Ok(reader)) => Ok(Self {
                _temporary_file: Some(temporary_file),
                writer: Box::new(writer),
                reader: Box::new(reader),
                entries: Vec::new(),
            }),
            (Err(error), _) | (_, Err(error)) => Err(error.into()),
        }
    }
}

fn box_disk_cache(cache: Result<DiskSharedStringCache>) -> Result<Box<dyn SharedStringCache>> {
    match cache {
        Ok(cache) => Ok(Box::new(cache)),
        Err(error) => Err(error),
    }
}

impl SharedStringCache for DiskSharedStringCache {
    fn put(&mut self, value: String) -> Result<()> {
        let offset = self.writer.seek(SeekFrom::End(0))?;
        let bytes = value.as_bytes();
        self.writer.write_all(bytes)?;
        self.entries.push((offset, bytes.len()));
        Ok(())
    }

    fn get(&mut self, index: usize) -> Result<String> {
        let (offset, length) = self.entries.get(index).copied().ok_or_else(|| {
            ExcelError::Format(format!("shared string index is out of bounds: {index}"))
        })?;
        self.reader.seek(SeekFrom::Start(offset))?;
        let mut bytes = vec![0; length];
        self.reader.read_exact(&mut bytes)?;
        String::from_utf8(bytes).map_err(|error| ExcelError::Format(error.to_string()))
    }

    #[cfg(test)]
    fn len(&self) -> usize {
        self.entries.len()
    }
}

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
    fn memory_and_disk_caches_preserve_indexed_utf8_values() {
        for mode in [ReadCacheMode::Memory, ReadCacheMode::Disk] {
            let mut cache = create_cache(mode, 0).expect("cache");
            assert_eq!(cache.len(), 0);
            cache.put("alpha".to_owned()).expect("first value");
            cache.put("中文😀".to_owned()).expect("Unicode value");
            assert_eq!(cache.len(), 2);
            assert_eq!(cache.get(0).expect("first value"), "alpha");
            assert_eq!(cache.get(1).expect("Unicode value"), "中文😀");
            assert!(cache.get(2).is_err());
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
    fn disk_cache_propagates_creation_seek_read_write_and_utf8_failures() {
        assert!(DiskSharedStringCache::from_temporary_file(Err(io_error())).is_err());
        let temporary_file = NamedTempFile::new().expect("temporary file");
        assert!(
            DiskSharedStringCache::from_handles(temporary_file, Err(io_error()), Err(io_error()),)
                .is_err()
        );
        let temporary_file = NamedTempFile::new().expect("temporary file");
        let writer = temporary_file.reopen().expect("writer");
        assert!(
            DiskSharedStringCache::from_handles(temporary_file, Ok(writer), Err(io_error()),)
                .is_err()
        );
        assert!(box_disk_cache(Err(ExcelError::Format("injected".to_owned()))).is_err());

        let fault = |fail_seek, fail_read, fail_write| FaultyIo {
            fail_seek,
            fail_read,
            fail_write,
        };
        let mut healthy = fault(false, false, false);
        let mut byte = [0_u8; 1];
        assert_eq!(healthy.read(&mut byte).expect("read"), 0);
        assert_eq!(healthy.write(&byte).expect("write"), 1);
        healthy.flush().expect("flush");
        let mut seek_write = DiskSharedStringCache {
            _temporary_file: None,
            writer: Box::new(fault(true, false, false)),
            reader: Box::new(fault(false, false, false)),
            entries: Vec::new(),
        };
        assert!(seek_write.put("value".to_owned()).is_err());
        let mut write = DiskSharedStringCache {
            _temporary_file: None,
            writer: Box::new(fault(false, false, true)),
            reader: Box::new(fault(false, false, false)),
            entries: Vec::new(),
        };
        assert!(write.put("value".to_owned()).is_err());
        let mut seek_read = DiskSharedStringCache {
            _temporary_file: None,
            writer: Box::new(fault(false, false, false)),
            reader: Box::new(fault(true, false, false)),
            entries: vec![(0, 1)],
        };
        assert!(seek_read.get(0).is_err());
        let mut read = DiskSharedStringCache {
            _temporary_file: None,
            writer: Box::new(fault(false, false, false)),
            reader: Box::new(fault(false, true, false)),
            entries: vec![(0, 1)],
        };
        assert!(read.get(0).is_err());

        let mut invalid_utf8 = DiskSharedStringCache::new().expect("disk cache");
        invalid_utf8.put("x".to_owned()).expect("cache value");
        invalid_utf8.writer.seek(SeekFrom::Start(0)).expect("seek");
        invalid_utf8
            .writer
            .write_all(&[0xff])
            .expect("corrupt value");
        assert!(invalid_utf8.get(0).is_err());
    }
}
