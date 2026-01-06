use napi::bindgen_prelude::*;
use napi_derive::napi;
use std::cell::RefCell;
use std::rc::Rc;

use crate::cell::CellValue;
use crate::workbook::Workbook as InnerWorkbook;

type CellInput = Either3<String, f64, bool>;

type SharedWorkbook = Rc<RefCell<InnerWorkbook>>;

#[napi]
pub struct Workbook {
    inner: SharedWorkbook,
}

#[napi]
impl Workbook {
    #[napi(constructor)]
    pub fn new() -> Self {
        Self {
            inner: Rc::new(RefCell::new(InnerWorkbook::new())),
        }
    }

    #[napi]
    pub fn add_worksheet(&self, name: String) -> Worksheet {
        let index = self.inner.borrow_mut().add_worksheet(name);
        Worksheet {
            workbook: Rc::clone(&self.inner),
            index,
        }
    }

    #[napi]
    pub fn save(&self, path: String) -> Result<()> {
        self.inner
            .borrow()
            .save(&path)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    #[napi(getter)]
    pub fn worksheet_count(&self) -> u32 {
        self.inner.borrow().worksheet_count() as u32
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

#[napi]
pub struct Worksheet {
    workbook: SharedWorkbook,
    index: usize,
}

#[napi]
impl Worksheet {
    #[napi(constructor)]
    pub fn new() -> Result<Self> {
        Err(Error::from_reason(
            "Worksheet cannot be constructed directly. Use Workbook.addWorksheet() instead.",
        ))
    }

    #[napi]
    pub fn write(&self, row: u32, col: u32, value: CellInput) -> Result<()> {
        let cell_value = match value {
            Either3::A(s) => CellValue::String(s),
            Either3::B(n) => CellValue::Number(n),
            Either3::C(b) => CellValue::Boolean(b),
        };
        self.workbook
            .borrow_mut()
            .write(self.index, row, col, cell_value)
            .map_err(|e| Error::from_reason(e.to_string()))
    }

    #[napi(getter)]
    pub fn name(&self) -> Result<String> {
        let workbook = self.workbook.borrow();
        let worksheet = workbook
            .get_worksheet(self.index)
            .ok_or_else(|| Error::from_reason("Worksheet not found"))?;
        Ok(worksheet.name().to_string())
    }
}
