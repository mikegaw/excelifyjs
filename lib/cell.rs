#[derive(Debug, Clone)]
pub enum CellValue {
    String(String),
    Number(f64),
    Boolean(bool),
    Empty,
}

impl CellValue {
    pub fn to_xlsx_value(&self) -> String {
        match self {
            CellValue::String(s) => s.clone(),
            CellValue::Number(n) => n.to_string(),
            CellValue::Boolean(b) => if *b { "1".to_string() } else { "0".to_string() },
            CellValue::Empty => String::new(),
        }
    }

    pub fn xlsx_type(&self) -> Option<&'static str> {
        match self {
            CellValue::String(_) => Some("inlineStr"),
            CellValue::Number(_) => None,
            CellValue::Boolean(_) => Some("b"),
            CellValue::Empty => None,
        }
    }
}

impl Default for CellValue {
    fn default() -> Self {
        CellValue::Empty
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::String(s)
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::String(s.to_string())
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

impl From<i64> for CellValue {
    fn from(n: i64) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
        CellValue::Number(n as f64)
    }
}

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        CellValue::Boolean(b)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_string_value() {
        let cell = CellValue::String("Hello".to_string());
        assert_eq!(cell.to_xlsx_value(), "Hello");
        assert_eq!(cell.xlsx_type(), Some("inlineStr"));
    }

    #[test]
    fn test_number_value() {
        let cell = CellValue::Number(42.5);
        assert_eq!(cell.to_xlsx_value(), "42.5");
        assert_eq!(cell.xlsx_type(), None);
    }

    #[test]
    fn test_boolean_value() {
        let cell_true = CellValue::Boolean(true);
        let cell_false = CellValue::Boolean(false);
        assert_eq!(cell_true.to_xlsx_value(), "1");
        assert_eq!(cell_false.to_xlsx_value(), "0");
        assert_eq!(cell_true.xlsx_type(), Some("b"));
    }

    #[test]
    fn test_from_conversions() {
        let _s: CellValue = "test".into();
        let _n: CellValue = 42.0_f64.into();
        let _b: CellValue = true.into();
    }
}
