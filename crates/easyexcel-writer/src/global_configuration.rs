//! Bridges [`GlobalConfiguration`] to [`WriteOptions`].

use easyexcel_core::metadata::GlobalConfiguration;

use crate::WriteOptions;

/// Builds a global configuration snapshot from write options.
///
/// Mirrors Java holder propagation from `WriteBasicParameter` into
/// `GlobalConfiguration`.
#[must_use]
pub fn global_configuration_from_write_options(options: &WriteOptions) -> GlobalConfiguration {
    GlobalConfiguration {
        auto_trim: options.auto_trim,
        use1904windowing: options.use_1904_windowing,
        locale: options.locale.clone(),
        use_scientific_format: options.use_scientific_format,
        filed_cache_location: options.filed_cache_location,
    }
}

/// Applies a global configuration onto write options without replacing sheet fields.
pub fn apply_global_configuration_to_write_options(
    global: &GlobalConfiguration,
    options: &mut WriteOptions,
) {
    options.auto_trim = global.auto_trim;
    options.use_1904_windowing = global.use1904windowing;
    options.locale = global.locale.clone();
    options.use_scientific_format = global.use_scientific_format;
    options.filed_cache_location = global.filed_cache_location;
}

#[cfg(test)]
mod tests {
    use easyexcel_core::CacheLocation;

    use super::*;

    #[test]
    fn write_bridge_keeps_defaults_without_overwriting_sheet_name() {
        let mut options = WriteOptions {
            sheet_name: "Custom".to_owned(),
            ..WriteOptions::default()
        };
        let config = global_configuration_from_write_options(&options);
        assert!(config.auto_trim());
        apply_global_configuration_to_write_options(&config, &mut options);
        assert_eq!(options.sheet_name, "Custom");
    }

    #[test]
    fn global_configuration_round_trips_write_options() {
        let mut options = WriteOptions::default();
        options.auto_trim = false;
        options.use_1904_windowing = true;
        options.use_scientific_format = true;
        options.locale = "zh_CN".to_owned();
        options.filed_cache_location = CacheLocation::Memory;

        let global = global_configuration_from_write_options(&options);
        let mut restored = WriteOptions::default();
        apply_global_configuration_to_write_options(&global, &mut restored);

        assert_eq!(restored.auto_trim, options.auto_trim);
        assert_eq!(restored.use_1904_windowing, options.use_1904_windowing);
        assert_eq!(
            restored.use_scientific_format,
            options.use_scientific_format
        );
        assert_eq!(restored.locale, options.locale);
        assert_eq!(restored.filed_cache_location, options.filed_cache_location);
    }
}
