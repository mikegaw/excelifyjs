use std::path::Path;

use crate::cell::CellValue;
use crate::error::{ExcelifyError, Result};
use crate::worksheet::Worksheet;
use crate::writer::XlsxWriter;

#[derive(Debug)]
pub struct Workbook {
    worksheets: Vec<Worksheet>,
}

impl Workbook {
    pub fn new() -> Self {
        Self {
            worksheets: Vec::new(),
        }
    }

    pub fn add_worksheet(&mut self, name: impl Into<String>) -> usize {
        let ws = Worksheet::new(name);
        self.worksheets.push(ws);
        self.worksheets.len() - 1
    }

    pub fn get_worksheet(&self, index: usize) -> Option<&Worksheet> {
        self.worksheets.get(index)
    }

    pub fn get_worksheet_mut(&mut self, index: usize) -> Option<&mut Worksheet> {
        self.worksheets.get_mut(index)
    }

    pub fn worksheets(&self) -> &[Worksheet] {
        &self.worksheets
    }

    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    pub fn write_string(
        &mut self,
        sheet_index: usize,
        row: u32,
        col: u32,
        value: impl Into<String>,
    ) -> Result<()> {
        let ws = self
            .worksheets
            .get_mut(sheet_index)
            .ok_or(ExcelifyError::SheetNotFound(sheet_index))?;
        ws.write_string(row, col, value);
        Ok(())
    }

    pub fn write_number(
        &mut self,
        sheet_index: usize,
        row: u32,
        col: u32,
        value: f64,
    ) -> Result<()> {
        let ws = self
            .worksheets
            .get_mut(sheet_index)
            .ok_or(ExcelifyError::SheetNotFound(sheet_index))?;
        ws.write_number(row, col, value);
        Ok(())
    }

    pub fn write_boolean(
        &mut self,
        sheet_index: usize,
        row: u32,
        col: u32,
        value: bool,
    ) -> Result<()> {
        let ws = self
            .worksheets
            .get_mut(sheet_index)
            .ok_or(ExcelifyError::SheetNotFound(sheet_index))?;
        ws.write_boolean(row, col, value);
        Ok(())
    }

    pub fn write(
        &mut self,
        sheet_index: usize,
        row: u32,
        col: u32,
        value: impl Into<CellValue>,
    ) -> Result<()> {
        let ws = self
            .worksheets
            .get_mut(sheet_index)
            .ok_or(ExcelifyError::SheetNotFound(sheet_index))?;
        ws.write(row, col, value);
        Ok(())
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let writer = XlsxWriter::new(self);
        writer.save(path)
    }
}

impl Default for Workbook {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_workbook() {
        let wb = Workbook::new();
        assert_eq!(wb.worksheet_count(), 0);
    }

    #[test]
    fn test_add_worksheet() {
        let mut wb = Workbook::new();
        let idx = wb.add_worksheet("Sheet1");
        assert_eq!(idx, 0);
        assert_eq!(wb.worksheet_count(), 1);
        assert_eq!(wb.get_worksheet(0).unwrap().name(), "Sheet1");
    }

    #[test]
    fn test_write_to_worksheet() {
        let mut wb = Workbook::new();
        wb.add_worksheet("Sheet1");

        wb.write_string(0, 0, 0, "Hello").unwrap();
        wb.write_number(0, 0, 1, 42.0).unwrap();
        wb.write_boolean(0, 1, 0, true).unwrap();

        let ws = wb.get_worksheet(0).unwrap();
        assert!(ws.get(0, 0).is_some());
        assert!(ws.get(0, 1).is_some());
        assert!(ws.get(1, 0).is_some());
    }

    #[test]
    fn test_write_to_invalid_sheet() {
        let mut wb = Workbook::new();
        let result = wb.write_string(0, 0, 0, "Hello");
        assert!(result.is_err());
    }
}
