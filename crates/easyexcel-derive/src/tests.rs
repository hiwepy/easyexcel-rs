use proc_macro_crate::FoundCrate;
use quote::quote;
use syn::{DeriveInput, parse_quote};

use super::*;

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
    let options = parse_struct_options(&input.attrs).expect("valid option");
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
        parse_struct_options(&input.attrs)
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
        assert!(parse_struct_options(&input.attrs).is_err());
    }
}

#[test]
fn field_options_parse_every_supported_value_and_reject_unknown_values() {
    let input: DeriveInput = parse_quote! {
        struct User {
            #[excel(name = "姓名", index = 2, order = 1, format = "%Y-%m-%d", converter = crate::NameConverter, column_width = 30, ignore)]
            name: String,
        }
    };
    let Data::Struct(data) = input.data else {
        panic!("expected struct");
    };
    let field = data.fields.iter().next().expect("field");
    let options = parse_field_options(&field.attrs).expect("valid options");
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

    let input: DeriveInput = parse_quote! {
        struct User { #[excel(unknown)] name: String }
    };
    let Data::Struct(data) = input.data else {
        panic!("expected struct");
    };
    let Err(error) = parse_field_options(&data.fields.iter().next().expect("field").attrs) else {
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
    ] {
        let source = format!("struct User {{ #[excel({attribute})] value: String }}");
        let input = syn::parse_str::<DeriveInput>(&source).expect("attribute tokens");
        let Data::Struct(data) = input.data else {
            panic!("expected struct");
        };
        assert!(
            parse_field_options(&data.fields.iter().next().expect("field").attrs).is_err(),
            "`{attribute}` must be rejected"
        );
    }
}

#[test]
fn expansion_generates_schema_readers_writers_defaults_and_generics() {
    let input: DeriveInput = parse_quote! {
        #[excel(ignore_unannotated, column_width = 25, head_row_height = 20, content_row_height = 16)]
        struct User<T>
        where
            T: Default,
        {
            #[excel(name = "姓名", index = 0, order = 2, format = "text", column_width = 30)]
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
        "ExcelWriteMetadata :: new () . column_width (25) . head_row_height (20) . content_row_height (16)",
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
