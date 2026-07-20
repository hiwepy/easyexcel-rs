//! Bridges [`GlobalConfiguration`] to [`ReadOptions`].

use easyexcel_core::metadata::GlobalConfiguration;

use crate::locale::ExcelLocale;
use crate::ScientificFormatMode;
use crate::ReadOptions;

/// Builds a global configuration snapshot from read options.
///
/// Mirrors Java holder propagation from `ReadBasicParameter` into
/// `GlobalConfiguration`.
#[must_use]
pub fn global_configuration_from_read_options(options: &ReadOptions) -> GlobalConfiguration {
    GlobalConfiguration {
        auto_trim: options.auto_trim,
        use1904windowing: options.use_1904_windowing,
        locale: options.locale.language_tag().to_owned(),
        use_scientific_format: matches!(
            options.scientific_format,
            ScientificFormatMode::Scientific
        ),
        filed_cache_location: easyexcel_core::CacheLocation::ThreadLocal,
    }
}

/// Applies a global configuration onto read options without replacing unrelated fields.
pub fn apply_global_configuration_to_read_options(
    global: &GlobalConfiguration,
    options: &mut ReadOptions,
) {
    options.auto_trim = global.auto_trim;
    options.use_1904_windowing = global.use1904windowing;
    if let Some(locale) = ExcelLocale::from_name(&global.locale) {
        options.locale = locale;
    }
    options.scientific_format = if global.use_scientific_format {
        ScientificFormatMode::Scientific
    } else {
        ScientificFormatMode::Plain
    };
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn global_configuration_round_trips_read_options() {
        let mut options = ReadOptions::default();
        options.auto_trim = false;
        options.use_1904_windowing = true;
        options.scientific_format = ScientificFormatMode::Scientific;

        let global = global_configuration_from_read_options(&options);
        let mut restored = ReadOptions::default();
        apply_global_configuration_to_read_options(&global, &mut restored);

        assert_eq!(restored.auto_trim, options.auto_trim);
        assert_eq!(restored.use_1904_windowing, options.use_1904_windowing);
        assert_eq!(restored.scientific_format, options.scientific_format);
    }
}
