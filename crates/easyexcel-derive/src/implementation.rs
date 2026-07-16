use proc_macro_crate::{FoundCrate, crate_name};
use quote::{format_ident, quote};
use syn::{Data, DeriveInput, Fields, LitInt, LitStr, Path};

#[derive(Default)]
struct FieldOptions {
    annotated: bool,
    ignore: bool,
    name: Option<LitStr>,
    index: Option<LitInt>,
    order: Option<LitInt>,
    format: Option<LitStr>,
}

pub(crate) fn expand_excel_row_tokens(
    input: proc_macro2::TokenStream,
) -> syn::Result<proc_macro2::TokenStream> {
    expand_excel_row(syn::parse2(input)?)
}

fn expand_excel_row(input: DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    let crate_path = easyexcel_path();
    let name = input.ident;
    let ignore_unannotated = parse_struct_options(&input.attrs)?;
    let fields = match input.data {
        Data::Struct(data) => match data.fields {
            Fields::Named(fields) => fields.named,
            _ => {
                return Err(syn::Error::new_spanned(
                    name,
                    "ExcelRow requires a struct with named fields",
                ));
            }
        },
        _ => {
            return Err(syn::Error::new_spanned(
                name,
                "ExcelRow can only be derived for structs",
            ));
        }
    };

    let mut columns = Vec::new();
    let mut readers = Vec::new();
    let mut writers = Vec::new();
    let mut schema_position = 0usize;

    for field in fields {
        let ident = field.ident.expect("named field");
        let ty = field.ty;
        let options = parse_field_options(&field.attrs)?;
        if options.ignore || (ignore_unannotated && !options.annotated) {
            readers.push(quote!(#ident: ::core::default::Default::default()));
            continue;
        }

        let field_name = ident.to_string();
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
        columns.push(quote!(
            #crate_path::ExcelColumn::new(
                #field_name,
                #header_name,
                #index,
                #order,
                #format,
            )
        ));
        let position = syn::Index::from(schema_position);
        readers.push(quote! {
            #ident: {
                let column = &Self::schema()[#position];
                let context = row.convert_context(column);
                <#ty as #crate_path::FromExcelCell>::from_excel_cell(
                    row.cell(column),
                    &context,
                )?
            }
        });
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
                #crate_path::IntoExcelCell::to_excel_cell(&self.#ident, &context)?
            }
        });
        schema_position += 1;
    }

    let (impl_generics, ty_generics, where_clause) = input.generics.split_for_impl();
    Ok(quote! {
        impl #impl_generics #crate_path::ExcelRow for #name #ty_generics #where_clause {
            fn schema() -> &'static [#crate_path::ExcelColumn] {
                const COLUMNS: &[#crate_path::ExcelColumn] = &[#(#columns),*];
                COLUMNS
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

fn parse_struct_options(attrs: &[syn::Attribute]) -> syn::Result<bool> {
    let mut ignore_unannotated = false;
    for attr in attrs.iter().filter(|attr| attr.path().is_ident("excel")) {
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("ignore_unannotated") {
                ignore_unannotated = true;
                Ok(())
            } else {
                Err(meta.error("unsupported ExcelRow struct option"))
            }
        })?;
    }
    Ok(ignore_unannotated)
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
            Err(meta.error("unsupported ExcelRow field option"))
        })?;
    }
    Ok(options)
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
