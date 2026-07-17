use proc_macro_crate::FoundCrate;
use quote::quote;
use syn::{DeriveInput, parse_quote};

use super::*;

fn assert_struct_style_options_rejected(attributes: &[&str]) {
    for attribute in attributes {
        let source = format!("#[excel({attribute})] struct User {{ value: String }}");
        let input = syn::parse_str::<DeriveInput>(&source).expect("attribute tokens");
        assert!(parse_struct_options(&input.attrs, &quote!(::easyexcel)).is_err());
    }
}

#[test]
fn token_entry_parses_valid_input_and_rejects_invalid_syntax() {
    assert!(
        expand_excel_row_tokens(quote!(
            struct User {
                value: String,
            }
        ))
        .expect("valid tokens")
        .to_string()
        .contains("ExcelRow")
    );
    assert!(expand_excel_row_tokens(quote!(struct)).is_err());
}

#[test]
fn crate_paths_support_self_renames_and_fallback_lookup() {
    assert_eq!(found_crate_path(FoundCrate::Itself).to_string(), "crate");
    assert_eq!(
        found_crate_path(FoundCrate::Name("easyexcel-renamed".to_owned())).to_string(),
        ":: easyexcel_renamed"
    );
    assert_eq!(resolve_easyexcel_path(None).to_string(), ":: easyexcel");
    assert_eq!(
        resolve_easyexcel_path(Some(FoundCrate::Name("renamed-core".to_owned()))).to_string(),
        ":: renamed_core"
    );
    assert!(!easyexcel_path().is_empty());
}

#[test]
fn struct_options_accept_ignore_unannotated_and_reject_unknown_values() {
    let input: DeriveInput = parse_quote! {
        #[excel(ignore_unannotated, column_width = 25, head_row_height = 20, content_row_height = 16)]
        struct User { name: String }
    };
    let options = parse_struct_options(&input.attrs, &quote!(::easyexcel)).expect("valid option");
    assert!(options.ignore_unannotated);
    assert_eq!(
        options
            .column_width
            .expect("width")
            .base10_parse::<u16>()
            .expect("u16"),
        25
    );
    assert_eq!(
        options
            .head_row_height
            .expect("head height")
            .base10_parse::<u16>()
            .expect("u16"),
        20
    );
    assert_eq!(
        options
            .content_row_height
            .expect("content height")
            .base10_parse::<u16>()
            .expect("u16"),
        16
    );

    let input: DeriveInput = parse_quote! {
        #[excel(unknown)]
        struct User { name: String }
    };
    assert!(
        parse_struct_options(&input.attrs, &quote!(::easyexcel))
            .err()
            .expect("unknown option")
            .to_string()
            .contains("unsupported ExcelRow struct option")
    );

    for attribute in [
        "column_width",
        "column_width = \"wide\"",
        "column_width = 65536",
        "head_row_height",
        "content_row_height = -1",
    ] {
        let source = format!("#[excel({attribute})] struct User {{ value: String }}");
        let input = syn::parse_str::<DeriveInput>(&source).expect("attribute tokens");
        assert!(parse_struct_options(&input.attrs, &quote!(::easyexcel)).is_err());
    }
}

#[test]
fn style_options_parse_java_equivalents_and_reject_invalid_values() {
    let input: DeriveInput = parse_quote! {
        #[excel(
            head_style(
                hidden = true,
                locked = false,
                quote_prefix = true,
                horizontal_alignment = "distributed",
                wrapped = true,
                vertical_alignment = "justify",
                rotation = 45,
                indent = 2,
                border_left = "thin",
                border_right = "medium",
                border_top = "dashed",
                border_bottom = "double",
                left_border_color = 0x112233,
                right_border_color = 0x223344,
                top_border_color = 0x334455,
                bottom_border_color = 0x445566,
                fill_pattern = "solid",
                fill_background_color = 0x556677,
                fill_foreground_color = 0x667788,
                shrink_to_fit = true,
                data_format = "0.00"
            ),
            content_style(wrapped = false, data_format = 14, fill_foreground_color = 10),
            head_font_style(
                font_name = "Arial",
                font_height_in_points = 12.5,
                italic = true,
                strikeout = false,
                color = 0x778899,
                type_offset = "superscript",
                underline = "double_accounting",
                charset = 1,
                bold = true
            ),
            content_font_style(bold = false)
        )]
        struct User { name: String }
    };
    let options = parse_struct_options(&input.attrs, &quote!(::easyexcel)).expect("valid styles");
    for style in [
        options.head_style,
        options.content_style,
        options.head_font_style,
        options.content_font_style,
    ] {
        assert!(style.expect("style").to_string().contains("style"));
    }

    assert_struct_style_options_rejected(&[
        "head_style(unknown = true)",
        "head_style(wrapped)",
        "head_style(wrapped = 1)",
        "head_style(rotation)",
        "head_style(rotation = \"up\")",
        "head_style(rotation = 32768)",
        "head_style(indent)",
        "head_style(indent = \"deep\")",
        "head_style(data_format)",
        "head_style(data_format = true)",
        "head_style(data_format = crate::FORMAT)",
        "head_style(data_format = 256)",
        "head_style(fill_foreground_color)",
        "head_style(fill_foreground_color = \"red\")",
        "head_style(horizontal_alignment)",
        "head_style(horizontal_alignment = 1)",
        "head_style(horizontal_alignment = \"diagonal\")",
        "head_style(vertical_alignment = \"diagonal\")",
        "head_style(border_left = \"triple\")",
        "head_style(fill_pattern = \"invalid\")",
        "head_style(indent = 256)",
        "head_style(left_border_color = 4294967296)",
        "head_style(foo::bar = true)",
        "head_font_style(unknown = true)",
        "head_font_style(font_name)",
        "head_font_style(font_name = 1)",
        "head_font_style(font_height_in_points)",
        "head_font_style(font_height_in_points = crate::SIZE)",
        "head_font_style(font_height_in_points = 1e999)",
        "head_font_style(bold)",
        "head_font_style(bold = 1)",
        "head_font_style(color)",
        "head_font_style(color = \"red\")",
        "head_font_style(charset)",
        "head_font_style(charset = \"default\")",
        "head_font_style(type_offset)",
        "head_font_style(type_offset = 1)",
        "head_font_style(font_height_in_points = \"large\")",
        "head_font_style(font_height_in_points = 0)",
        "head_font_style(charset = 256)",
        "head_font_style(type_offset = \"invalid\")",
        "head_font_style(underline = \"invalid\")",
        "head_font_style(foo::bar = true)",
    ]);

    assert_struct_style_options_rejected(&[
        "content_style(unknown = true)",
        "content_font_style(unknown = true)",
    ]);
}

#[test]
#[allow(clippy::too_many_lines)]
fn field_options_parse_every_supported_value_and_reject_unknown_values() {
    let input: DeriveInput = parse_quote! {
        struct User {
            #[excel(
                name = "姓名",
                index = 2,
                order = 1,
                format = "%Y-%m-%d",
                converter = crate::NameConverter,
                column_width = 30,
                head_style(wrapped = true),
                content_style(wrapped = false),
                head_font_style(bold = true),
                content_font_style(italic = true),
                ignore
            )]
            name: String,
        }
    };
    let Data::Struct(data) = input.data else {
        panic!("expected struct");
    };
    let field = data.fields.iter().next().expect("field");
    let options = parse_field_options(&field.attrs, &quote!(::easyexcel)).expect("valid options");
    assert!(options.annotated);
    assert!(options.ignore);
    assert_eq!(options.name.expect("name").value(), "姓名");
    assert_eq!(
        options
            .index
            .expect("index")
            .base10_parse::<usize>()
            .expect("usize"),
        2
    );
    assert_eq!(
        options
            .order
            .expect("order")
            .base10_parse::<i32>()
            .expect("i32"),
        1
    );
    assert_eq!(options.format.expect("format").value(), "%Y-%m-%d");
    assert_eq!(options.converter.expect("converter").segments.len(), 2);
    assert_eq!(
        options
            .column_width
            .expect("width")
            .base10_parse::<u16>()
            .expect("u16"),
        30
    );
    assert!(options.head_style.is_some());
    assert!(options.content_style.is_some());
    assert!(options.head_font_style.is_some());
    assert!(options.content_font_style.is_some());

    let input: DeriveInput = parse_quote! {
        struct User { #[excel(unknown)] name: String }
    };
    let Data::Struct(data) = input.data else {
        panic!("expected struct");
    };
    let Err(error) = parse_field_options(
        &data.fields.iter().next().expect("field").attrs,
        &quote!(::easyexcel),
    ) else {
        panic!("unknown option must be rejected");
    };
    assert!(
        error
            .to_string()
            .contains("unsupported ExcelRow field option")
    );

    for attribute in [
        "name",
        "name = 1",
        "index",
        "index = \"zero\"",
        "order",
        "order = \"first\"",
        "format",
        "format = 1",
        "converter",
        "converter = 1",
        "column_width",
        "column_width = \"wide\"",
        "column_width = 65536",
        "head_style(unknown = true)",
        "content_style(unknown = true)",
        "head_font_style(unknown = true)",
        "content_font_style(unknown = true)",
    ] {
        let source = format!("struct User {{ #[excel({attribute})] value: String }}");
        let input = syn::parse_str::<DeriveInput>(&source).expect("attribute tokens");
        let Data::Struct(data) = input.data else {
            panic!("expected struct");
        };
        assert!(
            parse_field_options(
                &data.fields.iter().next().expect("field").attrs,
                &quote!(::easyexcel),
            )
            .is_err(),
            "`{attribute}` must be rejected"
        );
    }
}

#[test]
fn expansion_generates_schema_readers_writers_defaults_and_generics() {
    let input: DeriveInput = parse_quote! {
        #[excel(
            ignore_unannotated,
            column_width = 25,
            head_row_height = 20,
            content_row_height = 16,
            head_style(fill_pattern = "solid"),
            content_style(wrapped = true),
            head_font_style(bold = true),
            content_font_style(italic = true)
        )]
        struct User<T>
        where
            T: Default,
        {
            #[excel(
                name = "姓名",
                index = 0,
                order = 2,
                format = "text",
                column_width = 30,
                head_style(wrapped = false),
                content_style(shrink_to_fit = true),
                head_font_style(font_name = "Arial"),
                content_font_style(bold = false)
            )]
            name: String,
            #[excel(ignore)]
            ignored: u32,
            unannotated: T,
        }
    };
    let expanded = expand_excel_row(input).expect("expansion").to_string();
    for expected in [
        "impl < T >",
        "ExcelRow for User < T >",
        "ExcelColumn :: new",
        "with_column_width (30)",
        "with_head_style",
        "with_content_style",
        "with_head_font_style",
        "with_content_font_style",
        "ExcelWriteMetadata :: new () . column_width (25) . head_row_height (20) . content_row_height (16) . head_style",
        "姓名",
        "Option :: Some (0)",
        "Option :: Some (\"text\")",
        "ignored : :: core :: default :: Default :: default ()",
        "unannotated : :: core :: default :: Default :: default ()",
        "FromExcelCell",
        "IntoExcelCell :: to_excel_cell",
    ] {
        assert!(
            expanded.contains(expected),
            "missing `{expected}` in {expanded}"
        );
    }

    let default_input: DeriveInput = parse_quote! {
        struct DefaultColumn { value: String }
    };
    let expanded = expand_excel_row(default_input)
        .expect("default expansion")
        .to_string();
    assert!(expanded.contains("\"value\""));
    assert!(expanded.contains("Option :: None"));
    assert!(expanded.contains("i32 :: MAX"));

    let converter_input: DeriveInput = parse_quote! {
        struct Converted {
            #[excel(converter = crate::NameConverter)]
            value: String,
        }
    };
    let expanded = expand_excel_row(converter_input)
        .expect("converter expansion")
        .to_string();
    for expected in [
        "Converter :: < String > :: convert_to_rust_data",
        "ReadConverterContext :: new",
        "NameConverter as :: core :: default :: Default",
        "Converter :: < String > :: convert_to_excel_data",
        "WriteConverterContext :: new",
    ] {
        assert!(
            expanded.contains(expected),
            "missing `{expected}` in {expanded}"
        );
    }
}

#[test]
fn expansion_rejects_tuple_structs_and_non_struct_items() {
    let tuple: DeriveInput = parse_quote!(
        struct Tuple(String);
    );
    assert!(
        expand_excel_row(tuple)
            .expect_err("tuple struct")
            .to_string()
            .contains("named fields")
    );

    let enumeration: DeriveInput = parse_quote!(
        enum Kind {
            One,
        }
    );
    assert!(
        expand_excel_row(enumeration)
            .expect_err("enum")
            .to_string()
            .contains("only be derived for structs")
    );

    let bad_struct_option: DeriveInput = parse_quote! {
        #[excel(unknown)]
        struct User { value: String }
    };
    assert!(expand_excel_row(bad_struct_option).is_err());

    let bad_field_option: DeriveInput = parse_quote! {
        struct User { #[excel(unknown)] value: String }
    };
    assert!(expand_excel_row(bad_field_option).is_err());
}

#[test]
fn generated_tokens_are_valid_rust_syntax() {
    let input: DeriveInput = parse_quote!(
        struct User {
            value: String,
        }
    );
    let tokens = expand_excel_row(input).expect("expansion");
    let wrapped = quote! { #tokens };
    assert!(!wrapped.is_empty());
}
