#![no_std]
extern crate alloc;

use console_error_panic_hook::set_once as set_panic_hook;
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
use web_sys::{BlobPropertyBag, HtmlElement};


#[wasm_bindgen(typescript_custom_section)]
const IGENERATE_OPTIONS: &'static str = r#"
export interface IGenerateOptions {
  failOnDuplicate?: boolean;
}
"#;

#[wasm_bindgen(typescript_custom_section)]
const IAPPEND_OPTIONS: &'static str = r#"
export interface IAppendOptions<T = any> {
  sheetName: string;
  data: T[];
  headers: Header[];
}

export interface Header {
  key: string;
  label: string;
  autoWidth?: boolean;
}
"#;

#[wasm_bindgen]
extern "C" {
  #[wasm_bindgen(typescript_type = "IAppendOptions")]
  pub type IAppendOptions;

  #[wasm_bindgen(typescript_type = "IGenerateOptions")]
  pub type IGenerateOptions;
}

#[derive(Serialize, Deserialize)]
pub struct Header {
  key: alloc::string::String,
  label: alloc::string::String,
  #[serde(rename(deserialize = "autoWidth"))]
  auto_width: Option<bool>,
}

#[derive(Serialize, Deserialize)]
pub struct AppendOptions {
  #[serde(rename(deserialize = "sheetName"))]
  sheet_name: Option<alloc::string::String>,
  data: alloc::vec::Vec<serde_json::Value>,
  headers: alloc::vec::Vec<Header>,
}

fn default_fail_on_duplicate() -> bool {
  true
}

#[derive(Serialize, Deserialize)]
struct GenerateOptions {
  #[serde(rename(deserialize = "failOnDuplicate"))]
  #[serde(default = "default_fail_on_duplicate")]
  /// If false, it will append a random string to the sheet name to make it unique if a sheet with the same name already exists.
  /// @default true
  fail_on_duplicate: bool,
}

impl Default for GenerateOptions {
  fn default() -> Self {
    Self {
      fail_on_duplicate: default_fail_on_duplicate(),
    }
  }
}

#[wasm_bindgen]
pub struct Converter {
  options: GenerateOptions,
  workbook: umya_spreadsheet::Spreadsheet,
}

#[wasm_bindgen]
impl Converter {
  #[wasm_bindgen(constructor)]
  pub fn new(options: Option<IGenerateOptions>) -> Self {
    set_panic_hook();

    let options: GenerateOptions = options
      .map(|options| serde_wasm_bindgen::from_value(options.into()).unwrap())
      .unwrap_or_default();

    let workbook = umya_spreadsheet::new_file_empty_worksheet();

    Self {
      options,
      workbook,
    }
  }

  pub fn append(&mut self, options: IAppendOptions) {
    let options: AppendOptions =
      serde_wasm_bindgen::from_value(options.into()).unwrap();
    let sheet_index = self.workbook.get_sheet_count();
    let mut sheet_name =
      options.sheet_name.unwrap_or(alloc::format!("Sheet{}", sheet_index + 1));

    if self.options.fail_on_duplicate {
      if self.workbook.get_sheet_by_name(&sheet_name).is_some() {
        panic!("Sheet with the same name already exists");
      }
    } else {
      let mut index = 1;
      while self.workbook.get_sheet_by_name(&sheet_name).is_some() {
        sheet_name = alloc::format!("{} ({})", sheet_name, index);
        index += 1;
      }
    }

    let sheet_name = sheet_name;

    let worksheet =
      self.workbook.new_sheet(sheet_name).expect("Failed to create new sheet");

    for (i, header) in options.headers.iter().enumerate() {
      let column = i as u32 + 1;
      let coordinates = (column, 1);
      if let Some(auto_width) = header.auto_width {
        if auto_width {
          worksheet
            .get_column_dimension_by_number_mut(&column)
            .set_auto_width(true);
        }
      }

      let cell = worksheet.get_cell_mut(coordinates);

      cell.set_value_string(&header.label);
    }

    for (row_index, row) in options.data.iter().enumerate() {
      for (col_index, header) in options.headers.iter().enumerate() {
        let coordinates = ((col_index as u32 + 1), (row_index as u32 + 2));
        let cell = worksheet.get_cell_mut(coordinates);

        let value = row.get(&header.key).unwrap_or(&serde_json::Value::Null);
        match value {
          serde_json::Value::String(value) => {
            cell.set_value_string(value);
          }
          serde_json::Value::Number(value) => {
            cell.set_value_number(value.as_f64().unwrap_or(0.0));
          }
          _ => {
            cell.set_value_string("");
          }
        }
      }
    }
  }

  pub async fn append_async(&mut self, options: IAppendOptions) {
    self.append(options);
  }

  #[wasm_bindgen(getter)]
  pub fn data(&self) -> alloc::vec::Vec<u8> {
    if self.workbook.get_sheet_count() == 0 {
      return alloc::vec::Vec::new();
    }

    let mut buffer: alloc::vec::Vec<u8> = alloc::vec::Vec::new();
    umya_spreadsheet::writer::xlsx::write_writer(&self.workbook, &mut buffer)
      .expect("Failed to write workbook");

    buffer
  }

  /// Save the file
  pub fn save(&self, name: alloc::string::String) {
    let buffer = self.data();

    let uint8_array = web_sys::js_sys::Uint8Array::from(buffer.as_slice());

    let array = web_sys::js_sys::Array::new();
    array.push(&uint8_array);

    let blob_properties = BlobPropertyBag::new();
    blob_properties.set_type(
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );

    let blob = web_sys::Blob::new_with_u8_array_sequence_and_options(
      &array,
      &blob_properties,
    )
    .expect("Failed to create blob");

    let url = web_sys::Url::create_object_url_with_blob(&blob)
      .expect("Failed to create object url");

    let window = web_sys::window().expect("Failed to get window");

    let document = window.document().expect("Failed to get document");
    let anchor = document.create_element("a").expect("Failed to create anchor");

    anchor.set_attribute("href", &url).expect("Failed to set href");

    let name = if name.ends_with(".xlsx") {
      name
    } else {
      alloc::format!("{}.xlsx", name)
    };

    anchor.set_attribute("download", &name).expect("Failed to set download");

    let body = document.body().expect("Failed to get body");
    body.append_child(&anchor).expect("Failed to append anchor");

    anchor
      .dyn_ref::<HtmlElement>()
      .expect("Failed to convert to HtmlElement")
      .click();

    web_sys::Url::revoke_object_url(&url).expect("Cannot revoke object url");

    body.remove_child(&anchor).expect("Failed to remove anchor");
  }
}
