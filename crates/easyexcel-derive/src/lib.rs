//! Derive macros for typed Excel row mapping.

use proc_macro::TokenStream;

mod implementation;

/// Derives static Excel column metadata and bidirectional row conversion.
#[proc_macro_derive(ExcelRow, attributes(excel))]
pub fn derive_excel_row(input: TokenStream) -> TokenStream {
    implementation::expand_excel_row_tokens(input.into())
        .unwrap_or_else(syn::Error::into_compile_error)
        .into()
}
