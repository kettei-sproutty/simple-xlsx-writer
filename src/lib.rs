#![no_std]
extern crate alloc;

use console_error_panic_hook::set_once as set_panic_hook;
use serde::{Deserialize, Serialize};
use wasm_bindgen::prelude::*;
use web_sys::{BlobPropertyBag, HtmlElement};

#[wasm_bindgen(typescript_custom_section)]
const IGENERATE_OPTIONS: &'static str = r#"
interface IGenerateOptions<T = any> {
  sheetName: string;
   data: T[];
  headers: Header[];
}

interface Header {
  key: string;
  label: string;
}
"#;

#[wasm_bindgen]
extern "C" {
  #[wasm_bindgen(typescript_type = "IGenerateOptions")]
  pub type IGenerateOptions;
}

#[derive(Serialize, Deserialize)]
pub struct Header {
  key: alloc::string::String,
  label: alloc::string::String,
}

#[derive(Serialize, Deserialize)]
pub struct ConverterOptions {
  #[serde(rename(deserialize = "sheetName"))]
  sheet_name: Option<alloc::string::String>,
  data: alloc::vec::Vec<serde_json::Value>,
  headers: alloc::vec::Vec<Header>,
}

#[wasm_bindgen]
pub struct Converter {
  workbook: umya_spreadsheet::Spreadsheet,
}

#[wasm_bindgen]
impl Converter {
  #[wasm_bindgen(constructor)]
  pub fn new() -> Self {
    set_panic_hook();
    let workbook = umya_spreadsheet::new_file_empty_worksheet();

    Self {
      workbook,
    }
  }

  pub fn append(&mut self, options: IGenerateOptions) {
    let options: ConverterOptions =
      serde_wasm_bindgen::from_value(options.into()).unwrap();
    let sheet_index = self.workbook.get_sheet_count();
    let sheet_name =
      options.sheet_name.unwrap_or(alloc::format!("Sheet{}", sheet_index + 1));

    let worksheet =
      self.workbook.new_sheet(sheet_name).expect("Failed to create new sheet");

    for (i, header) in options.headers.iter().enumerate() {
      let coordinates = ((i as u32 + 1), 1);
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

  pub async fn append_async(&mut self, options: IGenerateOptions) {
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
