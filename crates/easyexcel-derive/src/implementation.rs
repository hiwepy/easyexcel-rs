use proc_macro_crate::{FoundCrate, crate_name};
use quote::{format_ident, quote};
use syn::{Data, DeriveInput, Fields, Lit, LitBool, LitInt, LitStr, Path, meta::ParseNestedMeta};

#[derive(Default)]
struct StructOptions {
    ignore_unannotated: bool,
    column_width: Option<LitInt>,
    head_row_height: Option<LitInt>,
    content_row_height: Option<LitInt>,
    head_style: Option<proc_macro2::TokenStream>,
    content_style: Option<proc_macro2::TokenStream>,
    head_font_style: Option<proc_macro2::TokenStream>,
    content_font_style: Option<proc_macro2::TokenStream>,
    once_absolute_merge: Option<proc_macro2::TokenStream>,
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
    head_style: Option<proc_macro2::TokenStream>,
    content_style: Option<proc_macro2::TokenStream>,
    head_font_style: Option<proc_macro2::TokenStream>,
    content_font_style: Option<proc_macro2::TokenStream>,
    content_loop_merge: Option<proc_macro2::TokenStream>,
}

pub(crate) fn expand_excel_row_tokens(
    input: proc_macro2::TokenStream,
) -> syn::Result<proc_macro2::TokenStream> {
    expand_excel_row(syn::parse2(input)?)
}

#[allow(clippy::too_many_lines)]
fn expand_excel_row(input: DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    let crate_path = easyexcel_path();
    let name = input.ident;
    let struct_options = parse_struct_options(&input.attrs, &crate_path)?;
    let fields = named_fields(&name, input.data)?.named;

    let mut columns = Vec::new();
    let mut readers = Vec::new();
    let mut registered_readers = Vec::new();
    let mut writers = Vec::new();
    let mut registered_writers = Vec::new();
    let mut schema_position = 0usize;

    for field in fields {
        let ident = field.ident.expect("named field");
        let ty = field.ty;
        let options = parse_field_options(&field.attrs, &crate_path)?;
        if options.ignore || (struct_options.ignore_unannotated && !options.annotated) {
            readers.push(quote!(#ident: ::core::default::Default::default()));
            registered_readers.push(quote!(#ident: ::core::default::Default::default()));
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
        let column = decorate_column(
            column,
            options.column_width,
            options.head_style,
            options.content_style,
            options.head_font_style,
            options.content_font_style,
            options.content_loop_merge,
        );
        columns.push(column);
        let position = syn::Index::from(schema_position);
        let read_conversion = field_read_conversion(&crate_path, &ty, converter.as_ref());
        let registered_read_conversion =
            field_registered_read_conversion(&crate_path, &ty, converter.as_ref());
        readers.push(quote! {
            #ident: {
                let column = &Self::schema()[#position];
                let context = row.convert_context(column);
                #read_conversion
            }
        });
        registered_readers.push(quote! {
            #ident: {
                let column = &Self::schema()[#position];
                let context = row.convert_context(column);
                #registered_read_conversion
            }
        });
        let write_conversion = field_write_conversion(&crate_path, &ty, &ident, converter.as_ref());
        let registered_write_conversion =
            field_registered_write_conversion(&crate_path, &ty, &ident, converter.as_ref());
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
        registered_writers.push(quote! {
            {
                let column = &Self::schema()[#position];
                let context = #crate_path::ConvertContext {
                    sheet_name: ::std::string::String::new(),
                    row_index: 0,
                    column_index: column.index,
                    field: column.field,
                    format: column.format,
                };
                #registered_write_conversion
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

            fn from_row_with_converters(
                row: &#crate_path::RowData,
                converters: &#crate_path::ConverterRegistry,
            ) -> #crate_path::Result<Self> {
                Ok(Self { #(#registered_readers),* })
            }

            fn to_row(&self) -> #crate_path::Result<::std::vec::Vec<#crate_path::CellValue>> {
                Ok(::std::vec![#(#writers),*])
            }

            fn to_row_with_converters(
                &self,
                converters: &#crate_path::ConverterRegistry,
            ) -> #crate_path::Result<::std::vec::Vec<#crate_path::CellValue>> {
                Ok(::std::vec![#(#registered_writers),*])
            }
        }
    })
}

fn decorate_column(
    mut column: proc_macro2::TokenStream,
    width: Option<LitInt>,
    head_style: Option<proc_macro2::TokenStream>,
    content_style: Option<proc_macro2::TokenStream>,
    head_font_style: Option<proc_macro2::TokenStream>,
    content_font_style: Option<proc_macro2::TokenStream>,
    content_loop_merge: Option<proc_macro2::TokenStream>,
) -> proc_macro2::TokenStream {
    if let Some(width) = width {
        column = quote!(#column.with_column_width(#width));
    }
    if let Some(style) = head_style {
        column = quote!(#column.with_head_style(#style));
    }
    if let Some(style) = content_style {
        column = quote!(#column.with_content_style(#style));
    }
    if let Some(style) = head_font_style {
        column = quote!(#column.with_head_font_style(#style));
    }
    if let Some(style) = content_font_style {
        column = quote!(#column.with_content_font_style(#style));
    }
    if let Some(merge) = content_loop_merge {
        column = quote!(#column.with_loop_merge(#merge));
    }
    column
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
                    &#crate_path::ReadConverterContext::with_formula(
                        row.cell(column),
                        row.formula(column),
                        column,
                        &context,
                    ),
                )?
            }
        },
    )
}

fn field_registered_read_conversion(
    crate_path: &proc_macro2::TokenStream,
    ty: &syn::Type,
    converter: Option<&Path>,
) -> proc_macro2::TokenStream {
    converter.map_or_else(
        || {
            quote! {
                if let ::core::option::Option::Some(value) = converters.convert_to_rust_data::<#ty>(
                    &#crate_path::ReadConverterContext::with_formula(
                        row.cell(column),
                        row.formula(column),
                        column,
                        &context,
                    ),
                )? {
                    value
                } else {
                    <#ty as #crate_path::FromExcelCell>::from_excel_cell(
                        row.cell(column),
                        &context,
                    )?
                }
            }
        },
        |converter| {
            quote! {
                #crate_path::Converter::<#ty>::convert_to_rust_data(
                    &<#converter as ::core::default::Default>::default(),
                    &#crate_path::ReadConverterContext::with_formula(
                        row.cell(column),
                        row.formula(column),
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

fn field_registered_write_conversion(
    crate_path: &proc_macro2::TokenStream,
    ty: &syn::Type,
    ident: &syn::Ident,
    converter: Option<&Path>,
) -> proc_macro2::TokenStream {
    converter.map_or_else(
        || {
            quote! {
                if let ::core::option::Option::Some(value) = converters.convert_to_excel_data::<#ty>(
                    &self.#ident,
                    column,
                    &context,
                )? {
                    value
                } else {
                    #crate_path::IntoExcelCell::to_excel_cell(&self.#ident, &context)?
                }
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

fn parse_struct_options(
    attrs: &[syn::Attribute],
    crate_path: &proc_macro2::TokenStream,
) -> syn::Result<StructOptions> {
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
            if meta.path.is_ident("head_style") {
                options.head_style = Some(parse_cell_style(&meta, crate_path)?);
                return Ok(());
            }
            if meta.path.is_ident("content_style") {
                options.content_style = Some(parse_cell_style(&meta, crate_path)?);
                return Ok(());
            }
            if meta.path.is_ident("head_font_style") {
                options.head_font_style = Some(parse_font_style(&meta, crate_path)?);
                return Ok(());
            }
            if meta.path.is_ident("content_font_style") {
                options.content_font_style = Some(parse_font_style(&meta, crate_path)?);
                return Ok(());
            }
            if meta.path.is_ident("once_absolute_merge") {
                options.once_absolute_merge = Some(parse_once_absolute_merge(&meta, crate_path)?);
                return Ok(());
            }
            Err(meta.error("unsupported ExcelRow struct option"))
        })?;
    }
    Ok(options)
}

fn parse_field_options(
    attrs: &[syn::Attribute],
    crate_path: &proc_macro2::TokenStream,
) -> syn::Result<FieldOptions> {
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
            if meta.path.is_ident("head_style") {
                options.head_style = Some(parse_cell_style(&meta, crate_path)?);
                return Ok(());
            }
            if meta.path.is_ident("content_style") {
                options.content_style = Some(parse_cell_style(&meta, crate_path)?);
                return Ok(());
            }
            if meta.path.is_ident("head_font_style") {
                options.head_font_style = Some(parse_font_style(&meta, crate_path)?);
                return Ok(());
            }
            if meta.path.is_ident("content_font_style") {
                options.content_font_style = Some(parse_font_style(&meta, crate_path)?);
                return Ok(());
            }
            if meta.path.is_ident("content_loop_merge") {
                options.content_loop_merge = Some(parse_content_loop_merge(&meta, crate_path)?);
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

const HORIZONTAL_ALIGNMENT_VARIANTS: &[(&str, &str)] = &[
    ("general", "General"),
    ("left", "Left"),
    ("center", "Center"),
    ("right", "Right"),
    ("fill", "Fill"),
    ("justify", "Justify"),
    ("center_across", "CenterAcross"),
    ("distributed", "Distributed"),
];
const VERTICAL_ALIGNMENT_VARIANTS: &[(&str, &str)] = &[
    ("top", "Top"),
    ("center", "Center"),
    ("bottom", "Bottom"),
    ("justify", "Justify"),
    ("distributed", "Distributed"),
];
const BORDER_STYLE_VARIANTS: &[(&str, &str)] = &[
    ("none", "None"),
    ("thin", "Thin"),
    ("medium", "Medium"),
    ("dashed", "Dashed"),
    ("dotted", "Dotted"),
    ("thick", "Thick"),
    ("double", "Double"),
    ("hair", "Hair"),
    ("medium_dashed", "MediumDashed"),
    ("dash_dot", "DashDot"),
    ("medium_dash_dot", "MediumDashDot"),
    ("dash_dot_dot", "DashDotDot"),
    ("medium_dash_dot_dot", "MediumDashDotDot"),
    ("slant_dash_dot", "SlantDashDot"),
];
const FILL_PATTERN_VARIANTS: &[(&str, &str)] = &[
    ("none", "None"),
    ("solid", "Solid"),
    ("medium_gray", "MediumGray"),
    ("dark_gray", "DarkGray"),
    ("light_gray", "LightGray"),
    ("dark_horizontal", "DarkHorizontal"),
    ("dark_vertical", "DarkVertical"),
    ("dark_down", "DarkDown"),
    ("dark_up", "DarkUp"),
    ("dark_grid", "DarkGrid"),
    ("dark_trellis", "DarkTrellis"),
    ("light_horizontal", "LightHorizontal"),
    ("light_vertical", "LightVertical"),
    ("light_down", "LightDown"),
    ("light_up", "LightUp"),
    ("light_grid", "LightGrid"),
    ("light_trellis", "LightTrellis"),
    ("gray_125", "Gray125"),
    ("gray_0625", "Gray0625"),
];

fn parse_cell_style(
    meta: &ParseNestedMeta<'_>,
    crate_path: &proc_macro2::TokenStream,
) -> syn::Result<proc_macro2::TokenStream> {
    let mut assignments = Vec::new();
    meta.parse_nested_meta(|property| {
        let name = property
            .path
            .get_ident()
            .ok_or_else(|| property.error("style property must be an identifier"))?;
        if let Some(assignment) = parse_cell_style_scalar(&property, name, crate_path)? {
            assignments.push(assignment);
            return Ok(());
        }
        let field = format_ident!("{name}");
        match name.to_string().as_str() {
            "horizontal_alignment" => {
                let value = parse_named_variant(
                    &property,
                    crate_path,
                    "ExcelHorizontalAlignment",
                    HORIZONTAL_ALIGNMENT_VARIANTS,
                )?;
                assignments.push(
                    quote!(style.horizontal_alignment = ::core::option::Option::Some(#value);),
                );
            }
            "vertical_alignment" => {
                let value = parse_named_variant(
                    &property,
                    crate_path,
                    "ExcelVerticalAlignment",
                    VERTICAL_ALIGNMENT_VARIANTS,
                )?;
                assignments
                    .push(quote!(style.vertical_alignment = ::core::option::Option::Some(#value);));
            }
            "border_left" | "border_right" | "border_top" | "border_bottom" => {
                let value = parse_named_variant(
                    &property,
                    crate_path,
                    "ExcelBorderStyle",
                    BORDER_STYLE_VARIANTS,
                )?;
                assignments.push(quote!(style.#field = ::core::option::Option::Some(#value);));
            }
            "fill_pattern" => {
                let value = parse_named_variant(
                    &property,
                    crate_path,
                    "ExcelFillPattern",
                    FILL_PATTERN_VARIANTS,
                )?;
                assignments
                    .push(quote!(style.fill_pattern = ::core::option::Option::Some(#value);));
            }
            _ => return Err(property.error("unsupported cell style property")),
        }
        Ok(())
    })?;
    Ok(quote!({
        let mut style = #crate_path::ExcelCellStyle::new();
        #(#assignments)*
        style
    }))
}

fn parse_cell_style_scalar(
    property: &ParseNestedMeta<'_>,
    name: &syn::Ident,
    crate_path: &proc_macro2::TokenStream,
) -> syn::Result<Option<proc_macro2::TokenStream>> {
    let field = format_ident!("{name}");
    let assignment = match name.to_string().as_str() {
        "hidden" | "locked" | "quote_prefix" | "wrapped" | "shrink_to_fit" => {
            let value: LitBool = property.value()?.parse()?;
            quote!(style.#field = ::core::option::Option::Some(#value);)
        }
        "left_border_color"
        | "right_border_color"
        | "top_border_color"
        | "bottom_border_color"
        | "fill_background_color"
        | "fill_foreground_color" => {
            let value = parse_integer::<u32>(property)?;
            quote!(style.#field = ::core::option::Option::Some(
                #crate_path::ExcelColor::java_or_rgb(#value)
            );)
        }
        "rotation" => {
            let value = parse_integer::<i16>(property)?;
            quote!(style.rotation = ::core::option::Option::Some(#value);)
        }
        "indent" => {
            let value = parse_integer::<u8>(property)?;
            quote!(style.indent = ::core::option::Option::Some(#value);)
        }
        "data_format" => {
            let value: Lit = property.value()?.parse()?;
            match value {
                Lit::Str(value) => quote!(style.data_format = ::core::option::Option::Some(
                    #crate_path::ExcelDataFormat::Custom(#value)
                );),
                Lit::Int(value) => {
                    value
                        .base10_parse::<u8>()
                        .map_err(|error| syn::Error::new_spanned(&value, error))?;
                    quote!(style.data_format = ::core::option::Option::Some(
                        #crate_path::ExcelDataFormat::Builtin(#value)
                    );)
                }
                value => {
                    return Err(syn::Error::new_spanned(
                        value,
                        "data format must be a built-in index or custom format string",
                    ));
                }
            }
        }
        _ => return Ok(None),
    };
    Ok(Some(assignment))
}

fn parse_font_style(
    meta: &ParseNestedMeta<'_>,
    crate_path: &proc_macro2::TokenStream,
) -> syn::Result<proc_macro2::TokenStream> {
    let mut assignments = Vec::new();
    meta.parse_nested_meta(|property| {
        let name = property
            .path
            .get_ident()
            .ok_or_else(|| property.error("font property must be an identifier"))?;
        match name.to_string().as_str() {
            "font_name" => {
                let value: LitStr = property.value()?.parse()?;
                assignments.push(quote!(style.font_name = ::core::option::Option::Some(#value);));
            }
            "font_height_in_points" => {
                let value: Lit = property.value()?.parse()?;
                let numeric = match &value {
                    Lit::Int(value) => value.base10_parse::<f64>(),
                    Lit::Float(value) => value.base10_parse::<f64>(),
                    _ => {
                        return Err(syn::Error::new_spanned(
                            value,
                            "font height must be numeric",
                        ));
                    }
                }
                .unwrap_or(f64::NAN);
                if !numeric.is_finite() || numeric <= 0.0 {
                    return Err(syn::Error::new_spanned(
                        value,
                        "font height must be positive",
                    ));
                }
                assignments.push(
                    quote!(style.font_height_in_points = ::core::option::Option::Some(#numeric);),
                );
            }
            "italic" | "strikeout" | "bold" => {
                let field = format_ident!("{name}");
                let value: LitBool = property.value()?.parse()?;
                assignments.push(quote!(style.#field = ::core::option::Option::Some(#value);));
            }
            "color" => {
                let value = parse_integer::<u32>(&property)?;
                assignments.push(quote!(style.color = ::core::option::Option::Some(
                    #crate_path::ExcelColor::java_or_rgb(#value)
                );));
            }
            "charset" => {
                let value = parse_integer::<u8>(&property)?;
                assignments.push(quote!(style.charset = ::core::option::Option::Some(#value);));
            }
            "type_offset" => {
                let value = parse_named_variant(
                    &property,
                    crate_path,
                    "ExcelFontScript",
                    &[
                        ("none", "None"),
                        ("superscript", "Superscript"),
                        ("subscript", "Subscript"),
                    ],
                )?;
                assignments.push(quote!(style.type_offset = ::core::option::Option::Some(#value);));
            }
            "underline" => {
                let value = parse_named_variant(
                    &property,
                    crate_path,
                    "ExcelUnderline",
                    &[
                        ("none", "None"),
                        ("single", "Single"),
                        ("double", "Double"),
                        ("single_accounting", "SingleAccounting"),
                        ("double_accounting", "DoubleAccounting"),
                    ],
                )?;
                assignments.push(quote!(style.underline = ::core::option::Option::Some(#value);));
            }
            _ => return Err(property.error("unsupported font style property")),
        }
        Ok(())
    })?;
    Ok(quote!({
        let mut style = #crate_path::ExcelFontStyle::new();
        #(#assignments)*
        style
    }))
}

/// Parses `content_loop_merge(each_row = N, column_extend = M)` into
/// [`easyexcel_core::LoopMergeProperty`] construction tokens.
fn parse_content_loop_merge(
    meta: &ParseNestedMeta<'_>,
    crate_path: &proc_macro2::TokenStream,
) -> syn::Result<proc_macro2::TokenStream> {
    let mut each_row: Option<LitInt> = None;
    let mut column_extend: Option<LitInt> = None;
    meta.parse_nested_meta(|property| {
        if property.path.is_ident("each_row") {
            each_row = Some(parse_integer::<u32>(&property)?);
            return Ok(());
        }
        if property.path.is_ident("column_extend") {
            column_extend = Some(parse_integer::<u16>(&property)?);
            return Ok(());
        }
        Err(property.error("unsupported content_loop_merge property"))
    })?;
    // Java defaults: eachRow = 1, columnExtend = 1.
    let each_row = each_row.unwrap_or_else(|| LitInt::new("1", proc_macro2::Span::call_site()));
    let column_extend =
        column_extend.unwrap_or_else(|| LitInt::new("1", proc_macro2::Span::call_site()));
    Ok(quote!(#crate_path::LoopMergeProperty::new(#each_row, #column_extend)))
}

/// Parses `once_absolute_merge(first_row_index = ..., ...)` into
/// [`easyexcel_core::OnceAbsoluteMergeProperty`] construction tokens.
fn parse_once_absolute_merge(
    meta: &ParseNestedMeta<'_>,
    crate_path: &proc_macro2::TokenStream,
) -> syn::Result<proc_macro2::TokenStream> {
    let mut first_row_index: Option<LitInt> = None;
    let mut last_row_index: Option<LitInt> = None;
    let mut first_column_index: Option<LitInt> = None;
    let mut last_column_index: Option<LitInt> = None;
    meta.parse_nested_meta(|property| {
        if property.path.is_ident("first_row_index") {
            first_row_index = Some(parse_signed_integer(&property)?);
            return Ok(());
        }
        if property.path.is_ident("last_row_index") {
            last_row_index = Some(parse_signed_integer(&property)?);
            return Ok(());
        }
        if property.path.is_ident("first_column_index") {
            first_column_index = Some(parse_signed_integer(&property)?);
            return Ok(());
        }
        if property.path.is_ident("last_column_index") {
            last_column_index = Some(parse_signed_integer(&property)?);
            return Ok(());
        }
        Err(property.error("unsupported once_absolute_merge property"))
    })?;
    // Java defaults are -1 (unset).
    let default = || LitInt::new("-1", proc_macro2::Span::call_site());
    let first_row_index = first_row_index.unwrap_or_else(default);
    let last_row_index = last_row_index.unwrap_or_else(default);
    let first_column_index = first_column_index.unwrap_or_else(default);
    let last_column_index = last_column_index.unwrap_or_else(default);
    Ok(quote!(#crate_path::OnceAbsoluteMergeProperty::new(
        #first_row_index,
        #last_row_index,
        #first_column_index,
        #last_column_index,
    )))
}

fn parse_integer<T>(meta: &ParseNestedMeta<'_>) -> syn::Result<LitInt>
where
    T: std::str::FromStr,
    T::Err: std::fmt::Display,
{
    let value: LitInt = meta.value()?.parse()?;
    value
        .base10_parse::<T>()
        .map_err(|error| syn::Error::new_spanned(&value, error))?;
    Ok(value)
}

/// Parses a signed integer literal, including unary negation (`-1`).
fn parse_signed_integer(meta: &ParseNestedMeta<'_>) -> syn::Result<LitInt> {
    let expr: syn::Expr = meta.value()?.parse()?;
    match expr {
        syn::Expr::Lit(syn::ExprLit {
            lit: Lit::Int(value),
            ..
        }) => {
            value
                .base10_parse::<i32>()
                .map_err(|error| syn::Error::new_spanned(&value, error))?;
            Ok(value)
        }
        syn::Expr::Unary(syn::ExprUnary {
            op: syn::UnOp::Neg(_),
            expr,
            ..
        }) => match *expr {
            syn::Expr::Lit(syn::ExprLit {
                lit: Lit::Int(value),
                ..
            }) => {
                let magnitude = value
                    .base10_parse::<i32>()
                    .map_err(|error| syn::Error::new_spanned(&value, error))?;
                let negated = magnitude
                    .checked_neg()
                    .ok_or_else(|| syn::Error::new_spanned(&value, "integer overflow"))?;
                Ok(LitInt::new(&negated.to_string(), value.span()))
            }
            other => Err(syn::Error::new_spanned(
                other,
                "merge index must be an integer",
            )),
        },
        other => Err(syn::Error::new_spanned(
            other,
            "merge index must be an integer",
        )),
    }
}

fn parse_named_variant(
    meta: &ParseNestedMeta<'_>,
    crate_path: &proc_macro2::TokenStream,
    enum_name: &str,
    variants: &[(&str, &str)],
) -> syn::Result<proc_macro2::TokenStream> {
    let value: LitStr = meta.value()?.parse()?;
    let variant = variants
        .iter()
        .find_map(|(name, variant)| (*name == value.value()).then_some(*variant))
        .ok_or_else(|| syn::Error::new_spanned(&value, format!("unsupported {enum_name} value")))?;
    let enum_name = format_ident!("{enum_name}");
    let variant = format_ident!("{variant}");
    Ok(quote!(#crate_path::#enum_name::#variant))
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
    if let Some(style) = &options.head_style {
        metadata = quote!(#metadata.head_style(#style));
    }
    if let Some(style) = &options.content_style {
        metadata = quote!(#metadata.content_style(#style));
    }
    if let Some(style) = &options.head_font_style {
        metadata = quote!(#metadata.head_font_style(#style));
    }
    if let Some(style) = &options.content_font_style {
        metadata = quote!(#metadata.content_font_style(#style));
    }
    if let Some(merge) = &options.once_absolute_merge {
        metadata = quote!(#metadata.once_absolute_merge(#merge));
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
