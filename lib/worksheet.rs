use std::collections::HashMap;

use crate::cell::CellValue;

#[derive(Debug)]
pub struct Worksheet {
    name: String,
    cells: HashMap<(u32, u32), CellValue>,
    max_row: u32,
    max_col: u32,
}

impl Worksheet {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            cells: HashMap::new(),
            max_row: 0,
            max_col: 0,
        }
    }

    pub fn name(&self) -> &str {
        &self.name
    }

    pub fn write(&mut self, row: u32, col: u32, value: impl Into<CellValue>) {
        self.cells.insert((row, col), value.into());
        self.max_row = self.max_row.max(row);
        self.max_col = self.max_col.max(col);
    }

    pub fn write_string(&mut self, row: u32, col: u32, value: impl Into<String>) {
        self.write(row, col, CellValue::String(value.into()));
    }

    pub fn write_number(&mut self, row: u32, col: u32, value: f64) {
        self.write(row, col, CellValue::Number(value));
    }

    pub fn write_boolean(&mut self, row: u32, col: u32, value: bool) {
        self.write(row, col, CellValue::Boolean(value));
    }

    pub fn get(&self, row: u32, col: u32) -> Option<&CellValue> {
        self.cells.get(&(row, col))
    }

    pub fn cells(&self) -> &HashMap<(u32, u32), CellValue> {
        &self.cells
    }

    pub fn dimensions(&self) -> (u32, u32) {
        (self.max_row, self.max_col)
    }

    pub fn is_empty(&self) -> bool {
        self.cells.is_empty()
    }
}

pub fn col_to_letter(col: u32) -> String {
    let mut result = String::new();
    let mut n = col + 1;

    while n > 0 {
        n -= 1;
        result.insert(0, (b'A' + (n % 26) as u8) as char);
        n /= 26;
    }

    result
}

pub fn cell_reference(row: u32, col: u32) -> String {
    format!("{}{}", col_to_letter(col), row + 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_worksheet() {
        let ws = Worksheet::new("Sheet1");
        assert_eq!(ws.name(), "Sheet1");
        assert!(ws.is_empty());
    }

    #[test]
    fn test_write_and_get() {
        let mut ws = Worksheet::new("Test");
        ws.write_string(0, 0, "Hello");
        ws.write_number(0, 1, 42.0);
        ws.write_boolean(1, 0, true);

        assert!(matches!(ws.get(0, 0), Some(CellValue::String(_))));
        assert!(matches!(ws.get(0, 1), Some(CellValue::Number(_))));
        assert!(matches!(ws.get(1, 0), Some(CellValue::Boolean(_))));
        assert!(ws.get(1, 1).is_none());
    }

    #[test]
    fn test_dimensions() {
        let mut ws = Worksheet::new("Test");
        ws.write_string(5, 10, "value");
        assert_eq!(ws.dimensions(), (5, 10));
    }

    #[test]
    fn test_col_to_letter() {
        assert_eq!(col_to_letter(0), "A");
        assert_eq!(col_to_letter(1), "B");
        assert_eq!(col_to_letter(25), "Z");
        assert_eq!(col_to_letter(26), "AA");
        assert_eq!(col_to_letter(27), "AB");
        assert_eq!(col_to_letter(701), "ZZ");
        assert_eq!(col_to_letter(702), "AAA");
    }

    #[test]
    fn test_cell_reference() {
        assert_eq!(cell_reference(0, 0), "A1");
        assert_eq!(cell_reference(0, 1), "B1");
        assert_eq!(cell_reference(9, 2), "C10");
        assert_eq!(cell_reference(0, 26), "AA1");
    }
}
