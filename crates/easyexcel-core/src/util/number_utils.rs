//! Java-compatible number conversion helpers.
//!
//! Mirrors `com.alibaba.excel.util.NumberUtils` and the `DecimalFormat`
//! subset used by EasyExcel's built-in numeric string converters.

use std::str::FromStr;

use bigdecimal::BigDecimal;
use num_bigint::BigInt;

use crate::NumberRoundingMode;
use crate::excel_error::ExcelError;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum NonFiniteNumber {
    Nan,
    PositiveInfinity,
    NegativeInfinity,
}

#[derive(Debug, Clone)]
struct DecimalSubpattern {
    prefix: String,
    suffix: String,
    min_integer_digits: usize,
    min_fraction_digits: usize,
    max_fraction_digits: usize,
    grouping_size: Option<usize>,
    exponent_digits: Option<usize>,
    exponent_integer_digits: usize,
    multiplier: i32,
}

#[derive(Debug, Clone)]
struct DecimalPattern {
    positive: DecimalSubpattern,
    negative: DecimalSubpattern,
}

#[derive(Debug, Clone, Copy)]
struct PatternToken {
    value: char,
    literal: bool,
}

/// Formats a finite number like Java `NumberUtils.format`.
pub(crate) fn format_decimal(
    value: &BigDecimal,
    negative: bool,
    pattern: Option<&str>,
    rounding_mode: NumberRoundingMode,
) -> Result<String, ExcelError> {
    let Some(pattern) = pattern.filter(|pattern| !pattern.is_empty()) else {
        return Ok(value.to_plain_string());
    };
    let pattern = DecimalPattern::parse(pattern)?;
    pattern.format(value, negative, rounding_mode)
}

/// Formats Java `NaN` / infinity through DecimalFormat affixes.
pub(crate) fn format_non_finite(
    value: NonFiniteNumber,
    pattern: Option<&str>,
) -> Result<String, ExcelError> {
    if value == NonFiniteNumber::Nan {
        return Ok("NaN".to_owned());
    }
    let Some(pattern) = pattern.filter(|pattern| !pattern.is_empty()) else {
        return Ok(match value {
            NonFiniteNumber::PositiveInfinity => "Infinity",
            NonFiniteNumber::NegativeInfinity => "-Infinity",
            NonFiniteNumber::Nan => unreachable!(),
        }
        .to_owned());
    };
    let pattern = DecimalPattern::parse(pattern)?;
    let part = if value == NonFiniteNumber::NegativeInfinity {
        &pattern.negative
    } else {
        &pattern.positive
    };
    Ok(format!("{}∞{}", part.prefix, part.suffix))
}

/// Parses a string like Java `NumberUtils.parseBigDecimal`.
pub(crate) fn parse_decimal(value: &str, pattern: Option<&str>) -> Result<BigDecimal, ExcelError> {
    let Some(pattern) = pattern.filter(|pattern| !pattern.is_empty()) else {
        // Java `new BigDecimal(string)` does not trim leading/trailing spaces
        // and requires the complete input to be numeric.
        return BigDecimal::from_str(value)
            .map_err(|_| ExcelError::Format(format!("parseBigDecimal failed for {value:?}")));
    };
    DecimalPattern::parse(pattern)?.parse_number(value)
}

impl DecimalPattern {
    fn parse(pattern: &str) -> Result<Self, ExcelError> {
        let subpatterns = tokenize_pattern(pattern)?;
        let positive = DecimalSubpattern::parse(
            subpatterns
                .first()
                .ok_or_else(|| invalid_pattern(pattern, "missing positive subpattern"))?,
            pattern,
        )?;
        let negative = if let Some(tokens) = subpatterns.get(1) {
            DecimalSubpattern::parse(tokens, pattern)?
        } else {
            DecimalSubpattern {
                prefix: format!("-{}", positive.prefix),
                ..positive.clone()
            }
        };
        if subpatterns.len() > 2 {
            return Err(invalid_pattern(pattern, "more than two subpatterns"));
        }
        Ok(Self { positive, negative })
    }

    fn format(
        &self,
        value: &BigDecimal,
        negative: bool,
        rounding_mode: NumberRoundingMode,
    ) -> Result<String, ExcelError> {
        let part = if negative {
            &self.negative
        } else {
            &self.positive
        };
        let absolute = if value < &BigDecimal::from(0) {
            -value.clone()
        } else {
            value.clone()
        };
        let rounding_mode = if negative {
            match rounding_mode {
                NumberRoundingMode::Ceiling => NumberRoundingMode::Down,
                NumberRoundingMode::Floor => NumberRoundingMode::Up,
                other => other,
            }
        } else {
            rounding_mode
        };
        part.format_absolute(&absolute, rounding_mode)
    }

    fn parse_number(&self, input: &str) -> Result<BigDecimal, ExcelError> {
        // Java DecimalFormat tests the explicit/default negative prefix before
        // the positive one when the input starts with a minus sign.
        if let Some(result) = self.negative.parse_number(input, true)? {
            return Ok(result);
        }
        if let Some(result) = self.positive.parse_number(input, false)? {
            return Ok(result);
        }
        Err(ExcelError::Format(format!(
            "DecimalFormat could not parse {input:?}"
        )))
    }
}

impl DecimalSubpattern {
    fn parse(tokens: &[PatternToken], source: &str) -> Result<Self, ExcelError> {
        let first = tokens
            .iter()
            .position(is_numeric_pattern_token)
            .ok_or_else(|| invalid_pattern(source, "missing digit pattern"))?;
        let last = tokens
            .iter()
            .rposition(is_numeric_pattern_token)
            .expect("first numeric token exists");
        let prefix_tokens = &tokens[..first];
        let number_tokens = &tokens[first..=last];
        let suffix_tokens = &tokens[last + 1..];
        let prefix = render_affix(prefix_tokens);
        let suffix = render_affix(suffix_tokens);
        let multiplier = affix_multiplier(prefix_tokens, suffix_tokens)?;

        let exponent_index = number_tokens
            .iter()
            .position(|token| !token.literal && token.value == 'E');
        let (mantissa, exponent) = exponent_index.map_or((number_tokens, None), |index| {
            (&number_tokens[..index], Some(&number_tokens[index + 1..]))
        });
        let exponent_digits = exponent
            .map(|tokens| {
                if tokens.is_empty()
                    || tokens
                        .iter()
                        .any(|token| token.literal || token.value != '0')
                {
                    Err(invalid_pattern(source, "invalid exponent"))
                } else {
                    Ok(tokens.len())
                }
            })
            .transpose()?;

        let decimal_index = mantissa
            .iter()
            .position(|token| !token.literal && token.value == '.');
        let (integer, fraction) = decimal_index.map_or((mantissa, &[][..]), |index| {
            (&mantissa[..index], &mantissa[index + 1..])
        });
        if integer
            .iter()
            .any(|token| token.literal || !matches!(token.value, '#' | '0' | ','))
            || fraction
                .iter()
                .any(|token| token.literal || !matches!(token.value, '#' | '0'))
        {
            return Err(invalid_pattern(source, "invalid mantissa"));
        }
        let integer_digits = integer
            .iter()
            .filter(|token| matches!(token.value, '#' | '0'))
            .count();
        if integer_digits == 0 && fraction.is_empty() {
            return Err(invalid_pattern(source, "missing digit"));
        }
        let min_integer_digits = integer.iter().filter(|token| token.value == '0').count();
        let min_fraction_digits = fraction.iter().filter(|token| token.value == '0').count();
        let max_fraction_digits = fraction.len();
        let grouping_size = integer
            .iter()
            .rposition(|token| token.value == ',')
            .map(|index| {
                integer[index + 1..]
                    .iter()
                    .filter(|token| matches!(token.value, '#' | '0'))
                    .count()
            })
            .filter(|size| *size > 0);
        Ok(Self {
            prefix,
            suffix,
            min_integer_digits,
            min_fraction_digits,
            max_fraction_digits,
            grouping_size,
            exponent_digits,
            exponent_integer_digits: integer_digits.max(1),
            multiplier,
        })
    }

    fn format_absolute(
        &self,
        value: &BigDecimal,
        rounding_mode: NumberRoundingMode,
    ) -> Result<String, ExcelError> {
        let scaled = value * BigDecimal::from(self.multiplier);
        let number = if self.exponent_digits.is_some() {
            self.format_scientific(&scaled, rounding_mode)?
        } else {
            self.format_plain(&scaled, rounding_mode)?
        };
        Ok(format!("{}{}{}", self.prefix, number, self.suffix))
    }

    fn format_plain(
        &self,
        value: &BigDecimal,
        rounding_mode: NumberRoundingMode,
    ) -> Result<String, ExcelError> {
        let rounded = round_decimal(value, self.max_fraction_digits, rounding_mode)?;
        let mut text = rounded.to_plain_string();
        if let Some((integer, fraction)) = text.split_once('.') {
            let mut fraction = fraction.to_owned();
            while fraction.len() > self.min_fraction_digits && fraction.ends_with('0') {
                fraction.pop();
            }
            while fraction.len() < self.min_fraction_digits {
                fraction.push('0');
            }
            text = if fraction.is_empty() {
                integer.to_owned()
            } else {
                format!("{integer}.{fraction}")
            };
        } else if self.min_fraction_digits > 0 {
            text.push('.');
            text.push_str(&"0".repeat(self.min_fraction_digits));
        }
        let (integer, fraction) = text
            .split_once('.')
            .map_or((text.as_str(), None), |parts| (parts.0, Some(parts.1)));
        let mut integer = integer.to_owned();
        while integer.len() < self.min_integer_digits.max(1) {
            integer.insert(0, '0');
        }
        if let Some(grouping_size) = self.grouping_size {
            integer = group_integer(&integer, grouping_size);
        }
        Ok(fraction.map_or(integer.clone(), |fraction| format!("{integer}.{fraction}")))
    }

    fn format_scientific(
        &self,
        value: &BigDecimal,
        rounding_mode: NumberRoundingMode,
    ) -> Result<String, ExcelError> {
        let exponent_digits = self.exponent_digits.expect("scientific pattern");
        let (coefficient, scale) = value.as_bigint_and_exponent();
        let mut exponent = if coefficient == BigInt::from(0) {
            0
        } else {
            let digits = coefficient.to_str_radix(10).trim_start_matches('-').len() as i64;
            let scientific = digits - scale - 1;
            let width = self.exponent_integer_digits as i64;
            scientific.div_euclid(width) * width
        };
        let mut mantissa = BigDecimal::new(coefficient, scale + exponent);
        let mut formatted = self.format_plain(&mantissa, rounding_mode)?;
        let integer_digits = formatted.split('.').next().unwrap_or("").len();
        if integer_digits > self.exponent_integer_digits {
            exponent += self.exponent_integer_digits as i64;
            let (coefficient, scale) = value.as_bigint_and_exponent();
            mantissa = BigDecimal::new(coefficient, scale + exponent);
            formatted = self.format_plain(&mantissa, rounding_mode)?;
        }
        let sign = if exponent < 0 { "-" } else { "" };
        Ok(format!(
            "{formatted}E{sign}{:0width$}",
            exponent.unsigned_abs(),
            width = exponent_digits
        ))
    }

    fn parse_number(&self, input: &str, negative: bool) -> Result<Option<BigDecimal>, ExcelError> {
        let Some(mut remaining) = input.strip_prefix(&self.prefix) else {
            return Ok(None);
        };
        let mut byte_end = 0;
        let mut saw_digit = false;
        let mut saw_decimal = false;
        let mut saw_exponent = false;
        for (index, ch) in remaining.char_indices() {
            let accepted = if ch.is_ascii_digit() {
                saw_digit = true;
                true
            } else if ch == '.' && !saw_decimal && !saw_exponent {
                saw_decimal = true;
                true
            } else if ch == ',' && self.grouping_size.is_some() && !saw_decimal && !saw_exponent {
                true
            } else if matches!(ch, 'E' | 'e')
                && self.exponent_digits.is_some()
                && saw_digit
                && !saw_exponent
            {
                saw_exponent = true;
                true
            } else if matches!(ch, '+' | '-')
                && saw_exponent
                && remaining[..index].ends_with(['E', 'e'])
            {
                true
            } else {
                false
            };
            if !accepted {
                break;
            }
            byte_end = index + ch.len_utf8();
        }
        if !saw_digit {
            return Ok(None);
        }
        let numeric = &remaining[..byte_end];
        remaining = &remaining[byte_end..];
        if !self.suffix.is_empty() {
            let Some(after_suffix) = remaining.strip_prefix(&self.suffix) else {
                return Ok(None);
            };
            let _ = after_suffix;
        }
        let normalized = numeric.replace(',', "");
        let mut value = BigDecimal::from_str(&normalized)
            .map_err(|_| ExcelError::Format(format!("DecimalFormat could not parse {input:?}")))?;
        if self.multiplier != 1 {
            value /= self.multiplier;
        }
        if negative {
            value = -value;
        }
        Ok(Some(value))
    }
}

fn tokenize_pattern(pattern: &str) -> Result<Vec<Vec<PatternToken>>, ExcelError> {
    let mut subpatterns = vec![Vec::new()];
    let mut quoted = false;
    let mut chars = pattern.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '\'' {
            if chars.peek() == Some(&'\'') {
                chars.next();
                subpatterns
                    .last_mut()
                    .expect("one subpattern")
                    .push(PatternToken {
                        value: '\'',
                        literal: true,
                    });
            } else {
                quoted = !quoted;
            }
            continue;
        }
        if ch == ';' && !quoted {
            subpatterns.push(Vec::new());
            continue;
        }
        subpatterns
            .last_mut()
            .expect("one subpattern")
            .push(PatternToken {
                value: ch,
                literal: quoted,
            });
    }
    if quoted {
        return Err(invalid_pattern(pattern, "unterminated quote"));
    }
    Ok(subpatterns)
}

fn is_numeric_pattern_token(token: &PatternToken) -> bool {
    !token.literal && matches!(token.value, '#' | '0' | '.' | ',' | 'E')
}

fn render_affix(tokens: &[PatternToken]) -> String {
    tokens.iter().map(|token| token.value).collect()
}

fn affix_multiplier(prefix: &[PatternToken], suffix: &[PatternToken]) -> Result<i32, ExcelError> {
    let percent = prefix
        .iter()
        .chain(suffix)
        .filter(|token| !token.literal && token.value == '%')
        .count();
    let per_mille = prefix
        .iter()
        .chain(suffix)
        .filter(|token| !token.literal && token.value == '‰')
        .count();
    if percent + per_mille > 1 {
        return Err(ExcelError::Format(
            "DecimalFormat pattern contains multiple multipliers".to_owned(),
        ));
    }
    Ok(if percent == 1 {
        100
    } else if per_mille == 1 {
        1_000
    } else {
        1
    })
}

fn round_decimal(
    value: &BigDecimal,
    scale: usize,
    mode: NumberRoundingMode,
) -> Result<BigDecimal, ExcelError> {
    let scale = i64::try_from(scale)
        .map_err(|_| ExcelError::Format("DecimalFormat scale exceeds i64".to_owned()))?;
    if mode == NumberRoundingMode::Unnecessary {
        let truncated = value.with_scale_round(scale, bigdecimal::RoundingMode::Down);
        if &truncated != value {
            return Err(ExcelError::Format(
                "rounding necessary for RoundingMode.UNNECESSARY".to_owned(),
            ));
        }
        return Ok(truncated);
    }
    Ok(value.with_scale_round(scale, mode.bigdecimal().expect("non-UNNECESSARY mode")))
}

fn group_integer(value: &str, size: usize) -> String {
    let mut output = String::with_capacity(value.len() + value.len() / size);
    for (index, ch) in value.chars().enumerate() {
        if index > 0 && (value.len() - index).is_multiple_of(size) {
            output.push(',');
        }
        output.push(ch);
    }
    output
}

fn invalid_pattern(pattern: &str, reason: &str) -> ExcelError {
    ExcelError::Format(format!(
        "invalid DecimalFormat pattern {pattern:?}: {reason}"
    ))
}

/// Mirrors Java `NumberUtils.parseShort` without a format.
pub fn parse_short(value: &str) -> Result<i16, ExcelError> {
    parse_decimal(value, None).and_then(|value| decimal_java_i16(&value))
}

/// Mirrors Java `NumberUtils.parseLong` without a format.
pub fn parse_long(value: &str) -> Result<i64, ExcelError> {
    parse_decimal(value, None).and_then(|value| decimal_java_i64(&value))
}

/// Mirrors Java `NumberUtils.parseInteger` without a format.
pub fn parse_integer(value: &str) -> Result<i32, ExcelError> {
    parse_decimal(value, None).and_then(|value| decimal_java_i32(&value))
}

/// Mirrors Java `NumberUtils.parseFloat` without a format.
pub fn parse_float(value: &str) -> Result<f32, ExcelError> {
    parse_decimal(value, None).and_then(|value| {
        value
            .to_string()
            .parse()
            .map_err(|_| ExcelError::Format(format!("parseFloat failed for {value}")))
    })
}

/// Mirrors Java `NumberUtils.parseBigDecimal` without a format.
pub fn parse_big_decimal(value: &str) -> Result<BigDecimal, ExcelError> {
    parse_decimal(value, None)
}

/// Mirrors Java `NumberUtils.parseByte` without a format.
pub fn parse_byte(value: &str) -> Result<i8, ExcelError> {
    parse_decimal(value, None)
        .map(|value| i8::from_le_bytes(java_signed_low_bytes::<1>(&decimal_to_big_int(&value))))
}

/// Mirrors Java `NumberUtils.parseDouble` without a format.
pub fn parse_double(value: &str) -> Result<f64, ExcelError> {
    parse_decimal(value, None).and_then(|value| {
        value
            .to_string()
            .parse()
            .map_err(|_| ExcelError::Format(format!("parseDouble failed for {value}")))
    })
}

/// Mirrors Apache Commons `NumberUtils.createBigInteger`.
pub fn parse_big_int(value: &str) -> Result<BigInt, ExcelError> {
    BigInt::from_str(value)
        .map_err(|_| ExcelError::Format(format!("parseBigInteger failed for {value:?}")))
}

fn decimal_to_big_int(value: &BigDecimal) -> BigInt {
    value.with_scale(0).into_bigint_and_exponent().0
}

fn java_signed_low_bytes<const N: usize>(value: &BigInt) -> [u8; N] {
    let extension = if value.sign() == num_bigint::Sign::Minus {
        u8::MAX
    } else {
        0
    };
    let mut output = [extension; N];
    let source = value.to_signed_bytes_le();
    let count = source.len().min(N);
    output[..count].copy_from_slice(&source[..count]);
    output
}

fn decimal_java_i16(value: &BigDecimal) -> Result<i16, ExcelError> {
    Ok(i16::from_le_bytes(java_signed_low_bytes::<2>(
        &decimal_to_big_int(value),
    )))
}

fn decimal_java_i32(value: &BigDecimal) -> Result<i32, ExcelError> {
    Ok(i32::from_le_bytes(java_signed_low_bytes::<4>(
        &decimal_to_big_int(value),
    )))
}

fn decimal_java_i64(value: &BigDecimal) -> Result<i64, ExcelError> {
    Ok(i64::from_le_bytes(java_signed_low_bytes::<8>(
        &decimal_to_big_int(value),
    )))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn decimal(value: &str) -> BigDecimal {
        value.parse().unwrap()
    }

    #[test]
    fn decimal_format_matches_java_golden_patterns() {
        for (pattern, value, expected) in [
            ("#.##%", "1.235", "123.5%"),
            ("#", "1.235", "1"),
            ("0.00", "1.235", "1.24"),
            ("#,##0.00", "1234.5", "1,234.50"),
            ("0.00;[neg]0.00", "-1.235", "[neg]1.24"),
            ("0.00E00", "1235", "1.24E03"),
        ] {
            let value = decimal(value);
            assert_eq!(
                format_decimal(
                    &value,
                    value < BigDecimal::from(0),
                    Some(pattern),
                    NumberRoundingMode::HalfUp,
                )
                .unwrap(),
                expected
            );
        }
    }

    #[test]
    fn decimal_parse_matches_java_parse_position_behavior() {
        assert_eq!(
            parse_decimal("12.34%", Some("#.##%")).unwrap(),
            decimal("0.1234")
        );
        assert!(parse_decimal("12.34", Some("#.##%")).is_err());
        assert_eq!(
            parse_decimal("1,234.50", Some("#,##0.00")).unwrap(),
            decimal("1234.50")
        );
        assert_eq!(
            parse_decimal("1.00abc", Some("0.00")).unwrap(),
            decimal("1.00")
        );
        assert!(parse_decimal(" 1.00", Some("0.00")).is_err());
        assert!(parse_decimal("abc1.00", Some("0.00")).is_err());
    }

    #[test]
    fn no_format_is_full_input_big_decimal_and_unnecessary_rejects_rounding() {
        assert_eq!(parse_integer("1.00").unwrap(), 1);
        assert_eq!(parse_byte("255.9").unwrap(), -1);
        assert!(parse_big_decimal(" 1.00").is_err());
        assert!(parse_big_decimal("1.00 ").is_err());
        assert!(
            format_decimal(
                &decimal("1.001"),
                false,
                Some("0.00"),
                NumberRoundingMode::Unnecessary,
            )
            .is_err()
        );
    }

    #[test]
    fn all_java_rounding_modes_match_direction_and_tie_rules() {
        for (mode, positive, negative) in [
            (NumberRoundingMode::Up, "1.3", "-1.3"),
            (NumberRoundingMode::Down, "1.2", "-1.2"),
            (NumberRoundingMode::Ceiling, "1.3", "-1.2"),
            (NumberRoundingMode::Floor, "1.2", "-1.3"),
        ] {
            assert_eq!(
                format_decimal(&decimal("1.21"), false, Some("0.0"), mode).unwrap(),
                positive
            );
            assert_eq!(
                format_decimal(&decimal("-1.21"), true, Some("0.0"), mode).unwrap(),
                negative
            );
        }
        for (mode, expected) in [
            (NumberRoundingMode::HalfUp, "1.3"),
            (NumberRoundingMode::HalfDown, "1.2"),
            (NumberRoundingMode::HalfEven, "1.2"),
        ] {
            assert_eq!(
                format_decimal(&decimal("1.25"), false, Some("0.0"), mode).unwrap(),
                expected
            );
        }
    }

    #[test]
    fn quoted_affixes_per_mille_and_scientific_parse_are_supported() {
        assert_eq!(
            format_decimal(
                &decimal("12.5"),
                false,
                Some("'USD '0.00"),
                NumberRoundingMode::HalfUp,
            )
            .unwrap(),
            "USD 12.50"
        );
        assert_eq!(
            format_decimal(
                &decimal("0.01234"),
                false,
                Some("#.##‰"),
                NumberRoundingMode::HalfUp,
            )
            .unwrap(),
            "12.34‰"
        );
        assert_eq!(
            parse_decimal("1.24E03", Some("0.00E00")).unwrap(),
            decimal("1240")
        );
    }
}
