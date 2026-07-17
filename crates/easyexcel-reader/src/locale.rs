use std::str::FromStr;

use pure_rust_locales::Locale as SystemLocale;
use ssfmt::Locale as FormatLocale;

use crate::locale_generated::formatter_locale;

/// Locale data used by Java-compatible number and date formatting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExcelLocale {
    language_tag: String,
    formatter: FormatLocale,
}

impl ExcelLocale {
    /// Resolves a Java, POSIX, or BCP-47-style locale name.
    ///
    /// Examples include `en_US`, `zh-CN`, and `de_DE.UTF-8`.
    #[must_use]
    pub fn from_name(name: &str) -> Option<Self> {
        let normalized = normalize_locale_name(name);
        let locale = SystemLocale::from_str(&normalized).ok().or_else(|| {
            language_default(&normalized).and_then(|fallback| SystemLocale::from_str(fallback).ok())
        })?;
        Some(Self {
            language_tag: normalized,
            formatter: formatter_locale(locale),
        })
    }

    /// Returns the normalized locale name used to resolve formatting data.
    #[must_use]
    pub fn language_tag(&self) -> &str {
        &self.language_tag
    }

    pub(crate) fn formatter(&self) -> FormatLocale {
        self.formatter.clone()
    }

    fn system_default(locale: Option<&str>) -> Self {
        locale.and_then(Self::from_name).unwrap_or_else(Self::en_us)
    }

    fn en_us() -> Self {
        Self::from_name("en_US").expect("en_US locale data is always available")
    }
}

impl Default for ExcelLocale {
    fn default() -> Self {
        Self::system_default(sys_locale::get_locale().as_deref())
    }
}

fn normalize_locale_name(name: &str) -> String {
    let name = name.trim();
    if name.eq_ignore_ascii_case("c") || name.eq_ignore_ascii_case("posix") {
        return "POSIX".to_owned();
    }
    let (base, modifier) = match name.split_once('.') {
        Some((base, suffix)) => (base, suffix.split_once('@').map(|(_, modifier)| modifier)),
        None => match name.split_once('@') {
            Some((base, modifier)) => (base, Some(modifier)),
            None => (name, None),
        },
    };
    let mut normalized = base.replace('-', "_");
    if let Some(modifier) = modifier.filter(|value| !value.is_empty()) {
        normalized.push('@');
        normalized.push_str(modifier);
    }
    normalized
}

fn language_default(name: &str) -> Option<&'static str> {
    let language = name.split(['_', '@']).next().unwrap_or(name);
    match language.to_ascii_lowercase().as_str() {
        "ar" => Some("ar_SA"),
        "de" => Some("de_DE"),
        "en" => Some("en_US"),
        "es" => Some("es_ES"),
        "fr" => Some("fr_FR"),
        "hi" => Some("hi_IN"),
        "it" => Some("it_IT"),
        "ja" => Some("ja_JP"),
        "ko" => Some("ko_KR"),
        "pt" => Some("pt_BR"),
        "ru" => Some("ru_RU"),
        "zh" => Some("zh_CN"),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn locale_names_normalize_system_java_and_bcp47_forms() {
        assert_eq!(normalize_locale_name(" C "), "POSIX");
        assert_eq!(normalize_locale_name("posix"), "POSIX");
        assert_eq!(normalize_locale_name("zh-CN"), "zh_CN");
        assert_eq!(normalize_locale_name("de_DE.UTF-8"), "de_DE");
        assert_eq!(normalize_locale_name("sr_RS.UTF-8@latin"), "sr_RS@latin");
        assert_eq!(normalize_locale_name("sr_RS@latin"), "sr_RS@latin");
        assert_eq!(normalize_locale_name("sr_RS.UTF-8@"), "sr_RS");
    }

    #[test]
    fn locale_resolution_supports_full_names_language_defaults_and_fallback() {
        let german = ExcelLocale::from_name("de-DE").expect("German locale");
        assert_eq!(german.language_tag(), "de_DE");
        assert_eq!(german.formatter.decimal_separator, ',');
        assert_eq!(german.formatter.thousands_separator, '.');
        assert_eq!(german.formatter.month_names_full[0], "Januar");

        let chinese = ExcelLocale::from_name("zh").expect("Chinese language default");
        assert_eq!(chinese.language_tag(), "zh");
        assert_eq!(chinese.formatter.month_names_full[0], "一月");
        assert!(ExcelLocale::from_name("unknown_LOCALE").is_none());
        let posix = ExcelLocale::from_name("POSIX").expect("POSIX locale");
        assert_eq!(posix.formatter.thousands_separator, ',');

        assert_eq!(ExcelLocale::system_default(None).language_tag(), "en_US");
        assert_eq!(
            ExcelLocale::system_default(Some("fr_FR")).language_tag(),
            "fr_FR"
        );
        assert_eq!(
            ExcelLocale::system_default(Some("invalid")).language_tag(),
            "en_US"
        );
    }

    #[test]
    fn language_defaults_cover_java_common_locale_constants() {
        for (language, expected) in [
            ("ar", "ar_SA"),
            ("de", "de_DE"),
            ("en", "en_US"),
            ("es", "es_ES"),
            ("fr", "fr_FR"),
            ("hi", "hi_IN"),
            ("it", "it_IT"),
            ("ja", "ja_JP"),
            ("ko", "ko_KR"),
            ("pt", "pt_BR"),
            ("ru", "ru_RU"),
            ("zh", "zh_CN"),
        ] {
            assert_eq!(language_default(language), Some(expected));
        }
        assert_eq!(language_default(""), None);
        assert_eq!(language_default("xx"), None);
    }
}
