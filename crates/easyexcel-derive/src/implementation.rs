use proc_macro_crate::{FoundCrate, crate_name};
use quote::{format_ident, quote};
use syn::{Data, DeriveInput, Fields, LitInt, LitStr, Path, meta::ParseNestedMeta};

#[derive(Default)]
struct StructOptions {
    ignore_unannotated: bool,
    column_width: Option<LitInt>,
    head_row_height: Option<LitInt>,
    content_row_height: Option<LitInt>,
}

#[derive(Default)]
struct FieldOptions {
    annotated: bool,
    ignore: bool,
    name: Option<LitStr>,
    index: Option<LitInt>,
    order: Option<LitInt>,
    format: Option<LitStr>,
    converter: Option<Path>,
    column_width: Option<LitInt>,
}

pub(crate) fn expand_excel_row_tokens(
    input: proc_macro2::TokenStream,
) -> syn::Result<proc_macro2::TokenStream> {
    expand_excel_row(syn::parse2(input)?)
}

fn expand_excel_row(input: DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    let crate_path = easyexcel_path();
    let name = input.ident;
    let struct_options = parse_struct_options(&input.attrs)?;
    let fields = named_fields(&name, input.data)?.named;

    let mut columns = Vec::new();
    let mut readers = Vec::new();
    let mut writers = Vec::new();
    let mut schema_position = 0usize;

    for field in fields {
        let ident = field.ident.expect("named field");
        let ty = field.ty;
        let options = parse_field_options(&field.attrs)?;
        if options.ignore || (struct_options.ignore_unannotated && !options.annotated) {
            readers.push(quote!(#ident: ::core::default::Default::default()));
            continue;
        }

        let field_name = ident.to_string();
        let converter = options.converter;
        let header_name = options
            .name
            .unwrap_or_else(|| LitStr::new(&field_name, ident.span()));
        let index = options.index.map_or_else(
            || quote!(::core::option::Option::None),
            |value| quote!(::core::option::Option::Some(#value)),
        );
        let order = options
            .order
            .map_or_else(|| quote!(i32::MAX), |value| quote!(#value));
        let format = options.format.map_or_else(
            || quote!(::core::option::Option::None),
            |value| quote!(::core::option::Option::Some(#value)),
        );
        let column = quote!(
            #crate_path::ExcelColumn::new(
                #field_name,
                #header_name,
                #index,
                #order,
                #format,
            )
        );
        let column = options.column_width.map_or(
            column.clone(),
            |width| quote!(#column.with_column_width(#width)),
        );
        columns.push(column);
        let position = syn::Index::from(schema_position);
        let read_conversion = field_read_conversion(&crate_path, &ty, converter.as_ref());
        readers.push(quote! {
            #ident: {
                let column = &Self::schema()[#position];
                let context = row.convert_context(column);
                #read_conversion
            }
        });
        let write_conversion = field_write_conversion(&crate_path, &ty, &ident, converter.as_ref());
        writers.push(quote! {
            {
                let column = &Self::schema()[#position];
                let context = #crate_path::ConvertContext {
                    sheet_name: ::std::string::String::new(),
                    row_index: 0,
                    column_index: column.index,
                    field: column.field,
                    format: column.format,
                };
                #write_conversion
            }
        });
        schema_position += 1;
    }

    let (impl_generics, ty_generics, where_clause) = input.generics.split_for_impl();
    let write_metadata = write_metadata_tokens(&crate_path, &struct_options);
    Ok(quote! {
        impl #impl_generics #crate_path::ExcelRow for #name #ty_generics #where_clause {
            fn schema() -> &'static [#crate_path::ExcelColumn] {
                const COLUMNS: &[#crate_path::ExcelColumn] = &[#(#columns),*];
                COLUMNS
            }

            fn write_metadata() -> &'static #crate_path::ExcelWriteMetadata {
                const METADATA: #crate_path::ExcelWriteMetadata = #write_metadata;
                &METADATA
            }

            fn from_row(row: &#crate_path::RowData) -> #crate_path::Result<Self> {
                Ok(Self { #(#readers),* })
            }

            fn to_row(&self) -> #crate_path::Result<::std::vec::Vec<#crate_path::CellValue>> {
                Ok(::std::vec![#(#writers),*])
            }
        }
    })
}

fn named_fields(name: &syn::Ident, data: Data) -> syn::Result<syn::FieldsNamed> {
    match data {
        Data::Struct(data) => match data.fields {
            Fields::Named(fields) => Ok(fields),
            _ => Err(syn::Error::new_spanned(
                name,
                "ExcelRow requires a struct with named fields",
            )),
        },
        _ => Err(syn::Error::new_spanned(
            name,
            "ExcelRow can only be derived for structs",
        )),
    }
}

fn field_read_conversion(
    crate_path: &proc_macro2::TokenStream,
    ty: &syn::Type,
    converter: Option<&Path>,
) -> proc_macro2::TokenStream {
    converter.map_or_else(
        || {
            quote! {
                <#ty as #crate_path::FromExcelCell>::from_excel_cell(
                    row.cell(column),
                    &context,
                )?
            }
        },
        |converter| {
            quote! {
                #crate_path::Converter::<#ty>::convert_to_rust_data(
                    &<#converter as ::core::default::Default>::default(),
                    &#crate_path::ReadConverterContext::new(
                        row.cell(column),
                        column,
                        &context,
                    ),
                )?
            }
        },
    )
}

fn field_write_conversion(
    crate_path: &proc_macro2::TokenStream,
    ty: &syn::Type,
    ident: &syn::Ident,
    converter: Option<&Path>,
) -> proc_macro2::TokenStream {
    converter.map_or_else(
        || {
            quote! {
                #crate_path::IntoExcelCell::to_excel_cell(&self.#ident, &context)?
            }
        },
        |converter| {
            quote! {
                #crate_path::Converter::<#ty>::convert_to_excel_data(
                    &<#converter as ::core::default::Default>::default(),
                    &#crate_path::WriteConverterContext::new(
                        &self.#ident,
                        column,
                        &context,
                    ),
                )?
            }
        },
    )
}

fn parse_struct_options(attrs: &[syn::Attribute]) -> syn::Result<StructOptions> {
    let mut options = StructOptions::default();
    for attr in attrs.iter().filter(|attr| attr.path().is_ident("excel")) {
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("ignore_unannotated") {
                options.ignore_unannotated = true;
                return Ok(());
            }
            if meta.path.is_ident("column_width") {
                options.column_width = Some(parse_dimension(&meta)?);
                return Ok(());
            }
            if meta.path.is_ident("head_row_height") {
                options.head_row_height = Some(parse_dimension(&meta)?);
                return Ok(());
            }
            if meta.path.is_ident("content_row_height") {
                options.content_row_height = Some(parse_dimension(&meta)?);
                return Ok(());
            }
            Err(meta.error("unsupported ExcelRow struct option"))
        })?;
    }
    Ok(options)
}

fn parse_field_options(attrs: &[syn::Attribute]) -> syn::Result<FieldOptions> {
    let mut options = FieldOptions::default();
    for attr in attrs.iter().filter(|attr| attr.path().is_ident("excel")) {
        options.annotated = true;
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("ignore") {
                options.ignore = true;
                return Ok(());
            }
            if meta.path.is_ident("name") {
                options.name = Some(meta.value()?.parse()?);
                return Ok(());
            }
            if meta.path.is_ident("index") {
                options.index = Some(meta.value()?.parse()?);
                return Ok(());
            }
            if meta.path.is_ident("order") {
                options.order = Some(meta.value()?.parse()?);
                return Ok(());
            }
            if meta.path.is_ident("format") {
                options.format = Some(meta.value()?.parse()?);
                return Ok(());
            }
            if meta.path.is_ident("converter") {
                options.converter = Some(meta.value()?.parse()?);
                return Ok(());
            }
            if meta.path.is_ident("column_width") {
                options.column_width = Some(parse_dimension(&meta)?);
                return Ok(());
            }
            Err(meta.error("unsupported ExcelRow field option"))
        })?;
    }
    Ok(options)
}

fn parse_dimension(meta: &ParseNestedMeta<'_>) -> syn::Result<LitInt> {
    let value: LitInt = meta.value()?.parse()?;
    value
        .base10_parse::<u16>()
        .map_err(|error| syn::Error::new_spanned(&value, error))?;
    Ok(value)
}

fn write_metadata_tokens(
    crate_path: &proc_macro2::TokenStream,
    options: &StructOptions,
) -> proc_macro2::TokenStream {
    let mut metadata = quote!(#crate_path::ExcelWriteMetadata::new());
    if let Some(value) = &options.column_width {
        metadata = quote!(#metadata.column_width(#value));
    }
    if let Some(value) = &options.head_row_height {
        metadata = quote!(#metadata.head_row_height(#value));
    }
    if let Some(value) = &options.content_row_height {
        metadata = quote!(#metadata.content_row_height(#value));
    }
    metadata
}

fn easyexcel_path() -> proc_macro2::TokenStream {
    let found = ["easyexcel", "easyexcel-core"]
        .into_iter()
        .find_map(|package| crate_name(package).ok());
    resolve_easyexcel_path(found)
}

fn resolve_easyexcel_path(found: Option<FoundCrate>) -> proc_macro2::TokenStream {
    found.map_or_else(
        || {
            let fallback: Path = syn::parse_quote!(::easyexcel);
            quote!(#fallback)
        },
        found_crate_path,
    )
}

fn found_crate_path(found: FoundCrate) -> proc_macro2::TokenStream {
    match found {
        FoundCrate::Itself => quote!(crate),
        FoundCrate::Name(name) => {
            let ident = format_ident!("{}", name.replace('-', "_"));
            quote!(::#ident)
        }
    }
}

#[cfg(test)]
#[path = "tests.rs"]
mod tests;
