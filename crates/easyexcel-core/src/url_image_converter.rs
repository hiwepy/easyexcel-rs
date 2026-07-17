//! Mirrors Java `com.alibaba.excel.converters.url.UrlImageConverter` with
//! Java's default timeout values (1s connect, 5s read).

use std::fmt::Display;
use std::io::Read;
use std::time::Duration;

use ureq::Agent;
use url::Url;

use crate::cell_value::CellValue;
use crate::convert_context::ConvertContext;
use crate::converter::Converter;
use crate::excel_error::ExcelError;
use crate::into_excel_cell::IntoExcelCell;
use crate::write_converter_context::WriteConverterContext;

/// Java `UrlImageConverter` equivalent with Java's default timeout values.
///
/// Uses the `ureq` crate for HTTP; defaulting to 1s connect and 5s read
/// matches Java EasyExcel.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UrlImageConverter {
    connect_timeout: Duration,
    read_timeout: Duration,
}

impl UrlImageConverter {
    /// Java `EasyExcel`'s default URL connection timeout.
    pub const DEFAULT_CONNECT_TIMEOUT: Duration = Duration::from_secs(1);
    /// Java `EasyExcel`'s default URL response-read timeout.
    pub const DEFAULT_READ_TIMEOUT: Duration = Duration::from_secs(5);

    /// Creates a converter with explicit connection and response-read timeouts.
    #[must_use]
    pub const fn new(connect_timeout: Duration, read_timeout: Duration) -> Self {
        Self {
            connect_timeout,
            read_timeout,
        }
    }

    /// Returns the configured connection timeout. (Java `getConnectTimeout()`)
    #[must_use]
    pub const fn connect_timeout(self) -> Duration {
        self.connect_timeout
    }

    /// Returns the configured response-read timeout. (Java `getReadTimeout()`)
    #[must_use]
    pub const fn read_timeout(self) -> Duration {
        self.read_timeout
    }

    fn download(self, value: &Url) -> Result<Vec<u8>, ExcelError> {
        let agent: Agent = ureq::Agent::config_builder()
            .timeout_connect(Some(self.connect_timeout))
            .timeout_recv_body(Some(self.read_timeout))
            .build()
            .into();
        let mut response = agent.get(value.as_str()).call().map_err(url_image_error)?;
        let mut bytes = Vec::new();
        response
            .body_mut()
            .as_reader()
            .read_to_end(&mut bytes)
            .map_err(url_image_error)?;
        Ok(bytes)
    }
}

impl Default for UrlImageConverter {
    fn default() -> Self {
        Self::new(Self::DEFAULT_CONNECT_TIMEOUT, Self::DEFAULT_READ_TIMEOUT)
    }
}

impl Converter<Url> for UrlImageConverter {
    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, Url>,
    ) -> Result<CellValue, ExcelError> {
        self.download(context.value()).map(CellValue::Image)
    }
}

impl IntoExcelCell for Url {
    fn to_excel_cell(&self, _context: &ConvertContext) -> Result<CellValue, ExcelError> {
        UrlImageConverter::default()
            .download(self)
            .map(CellValue::Image)
    }
}

fn url_image_error(error: impl Display) -> ExcelError {
    ExcelError::Io(std::io::Error::other(error.to_string()))
}
