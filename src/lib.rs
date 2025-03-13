use wasm_bindgen::prelude::*;
use serde::{Deserialize, Serialize};
use std::io::Cursor;

#[global_allocator]
static ALLOC: wee_alloc::WeeAlloc = wee_alloc::WeeAlloc::INIT;

// Define a header structure.
#[derive(Serialize, Deserialize)]
pub struct Header {
  /// The label of the header, which will be displayed in the Excel file.
  pub label: String,
  /// The key of the header, which will be used to extract the data from the data object.
  pub key: String,
}

/// Options for generating an XLSX.
#[derive(Serialize, Deserialize)]
pub struct GenerateOptions {
  /// The name of the sheet in the Excel file, optional.
  pub sheet_name: Option<String>,
  /// The headers of the Excel file.
  pub headers: Vec<Header>,
  /// The data of the Excel file.
  pub data: Vec<serde_json::Value>,
}

#[wasm_bindgen]
pub struct Converter {
  workbook: umya_spreadsheet::Spreadsheet,
}

#[wasm_bindgen]
impl Converter {
  #[wasm_bindgen(constructor)]
  pub fn new() -> Self {
    let workbook = umya_spreadsheet::new_file_empty_worksheet();

    Self {
      workbook,
    }
  }

  #[wasm_bindgen]
  pub fn append(&mut self, options: JsValue) {
    let options: GenerateOptions = serde_wasm_bindgen::from_value(options).expect("Failed to parse options");

    let sheet_index = self.workbook.get_sheet_count();
    let sheet_name = options.sheet_name.unwrap_or(format!("Sheet{}", sheet_index + 1));
    
    let worksheet = self.workbook.new_sheet(sheet_name).expect("Failed to create new sheet");

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
          },
          serde_json::Value::Number(value) => {
            cell.set_value_number(value.as_f64().unwrap_or(0.0));
          },
          _ => {
            cell.set_value_string("");
          }
        }
      }
    }
  }

  #[wasm_bindgen(getter)]
  pub fn data(&self) -> Vec<u8> {
    if self.workbook.get_sheet_count() == 0 {
      return Vec::new();
    }

    let mut buffer: Cursor<Vec<u8>> = Cursor::new(Vec::new());
    umya_spreadsheet::writer::xlsx::write_writer(&self.workbook, &mut buffer).expect("Failed to write workbook");
    buffer.into_inner()
  }
}
