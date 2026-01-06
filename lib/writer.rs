use std::collections::BTreeMap;
use std::fs::File;
use std::io::{Cursor, Write};
use std::path::Path;

use quick_xml::events::{BytesDecl, BytesEnd, BytesStart, BytesText, Event};
use quick_xml::Writer;
use zip::write::FileOptions;
use zip::ZipWriter;

use crate::cell::CellValue;
use crate::error::Result;
use crate::workbook::Workbook;
use crate::worksheet::{cell_reference, Worksheet};

pub struct XlsxWriter<'a> {
    workbook: &'a Workbook,
}

impl<'a> XlsxWriter<'a> {
    pub fn new(workbook: &'a Workbook) -> Self {
        Self { workbook }
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let file = File::create(path)?;
        let mut zip = ZipWriter::new(file);
        let options: FileOptions = FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .compression_level(Some(6));

        self.write_content_types(&mut zip, options)?;
        self.write_rels(&mut zip, options)?;
        self.write_workbook_xml(&mut zip, options)?;
        self.write_workbook_rels(&mut zip, options)?;

        for (idx, worksheet) in self.workbook.worksheets().iter().enumerate() {
            self.write_worksheet_xml(&mut zip, options, idx, worksheet)?;
        }

        zip.finish()?;
        Ok(())
    }

    fn write_content_types(
        &self,
        zip: &mut ZipWriter<File>,
        options: FileOptions,
    ) -> Result<()> {
        zip.start_file("[Content_Types].xml", options)?;

        let mut writer = Writer::new(Cursor::new(Vec::new()));
        writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut types = BytesStart::new("Types");
        types.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/content-types",
        ));
        writer.write_event(Event::Start(types))?;

        let mut default_rels = BytesStart::new("Default");
        default_rels.push_attribute(("Extension", "rels"));
        default_rels.push_attribute((
            "ContentType",
            "application/vnd.openxmlformats-package.relationships+xml",
        ));
        writer.write_event(Event::Empty(default_rels))?;

        let mut default_xml = BytesStart::new("Default");
        default_xml.push_attribute(("Extension", "xml"));
        default_xml.push_attribute(("ContentType", "application/xml"));
        writer.write_event(Event::Empty(default_xml))?;

        let mut override_wb = BytesStart::new("Override");
        override_wb.push_attribute(("PartName", "/xl/workbook.xml"));
        override_wb.push_attribute((
            "ContentType",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        ));
        writer.write_event(Event::Empty(override_wb))?;

        for idx in 0..self.workbook.worksheet_count() {
            let mut override_sheet = BytesStart::new("Override");
            override_sheet.push_attribute(("PartName", format!("/xl/worksheets/sheet{}.xml", idx + 1).as_str()));
            override_sheet.push_attribute((
                "ContentType",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            ));
            writer.write_event(Event::Empty(override_sheet))?;
        }

        writer.write_event(Event::End(BytesEnd::new("Types")))?;

        zip.write_all(writer.into_inner().into_inner().as_slice())?;
        Ok(())
    }

    fn write_rels(&self, zip: &mut ZipWriter<File>, options: FileOptions) -> Result<()> {
        zip.start_file("_rels/.rels", options)?;

        let mut writer = Writer::new(Cursor::new(Vec::new()));
        writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut rels = BytesStart::new("Relationships");
        rels.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        ));
        writer.write_event(Event::Start(rels))?;

        let mut rel = BytesStart::new("Relationship");
        rel.push_attribute(("Id", "rId1"));
        rel.push_attribute((
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        ));
        rel.push_attribute(("Target", "xl/workbook.xml"));
        writer.write_event(Event::Empty(rel))?;

        writer.write_event(Event::End(BytesEnd::new("Relationships")))?;

        zip.write_all(writer.into_inner().into_inner().as_slice())?;
        Ok(())
    }

    fn write_workbook_xml(
        &self,
        zip: &mut ZipWriter<File>,
        options: FileOptions,
    ) -> Result<()> {
        zip.start_file("xl/workbook.xml", options)?;

        let mut writer = Writer::new(Cursor::new(Vec::new()));
        writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut workbook = BytesStart::new("workbook");
        workbook.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ));
        workbook.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(workbook))?;

        writer.write_event(Event::Start(BytesStart::new("sheets")))?;

        for (idx, worksheet) in self.workbook.worksheets().iter().enumerate() {
            let mut sheet = BytesStart::new("sheet");
            sheet.push_attribute(("name", worksheet.name()));
            sheet.push_attribute(("sheetId", (idx + 1).to_string().as_str()));
            sheet.push_attribute(("r:id", format!("rId{}", idx + 1).as_str()));
            writer.write_event(Event::Empty(sheet))?;
        }

        writer.write_event(Event::End(BytesEnd::new("sheets")))?;
        writer.write_event(Event::End(BytesEnd::new("workbook")))?;

        zip.write_all(writer.into_inner().into_inner().as_slice())?;
        Ok(())
    }

    fn write_workbook_rels(
        &self,
        zip: &mut ZipWriter<File>,
        options: FileOptions,
    ) -> Result<()> {
        zip.start_file("xl/_rels/workbook.xml.rels", options)?;

        let mut writer = Writer::new(Cursor::new(Vec::new()));
        writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut rels = BytesStart::new("Relationships");
        rels.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/package/2006/relationships",
        ));
        writer.write_event(Event::Start(rels))?;

        for idx in 0..self.workbook.worksheet_count() {
            let mut rel = BytesStart::new("Relationship");
            rel.push_attribute(("Id", format!("rId{}", idx + 1).as_str()));
            rel.push_attribute((
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            ));
            rel.push_attribute(("Target", format!("worksheets/sheet{}.xml", idx + 1).as_str()));
            writer.write_event(Event::Empty(rel))?;
        }

        writer.write_event(Event::End(BytesEnd::new("Relationships")))?;

        zip.write_all(writer.into_inner().into_inner().as_slice())?;
        Ok(())
    }

    fn write_worksheet_xml(
        &self,
        zip: &mut ZipWriter<File>,
        options: FileOptions,
        idx: usize,
        worksheet: &Worksheet,
    ) -> Result<()> {
        zip.start_file(format!("xl/worksheets/sheet{}.xml", idx + 1), options)?;

        let mut writer = Writer::new(Cursor::new(Vec::new()));
        writer.write_event(Event::Decl(BytesDecl::new("1.0", Some("UTF-8"), Some("yes"))))?;

        let mut ws = BytesStart::new("worksheet");
        ws.push_attribute((
            "xmlns",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ));
        ws.push_attribute((
            "xmlns:r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ));
        writer.write_event(Event::Start(ws))?;

        writer.write_event(Event::Start(BytesStart::new("sheetData")))?;

        let cells = worksheet.cells();
        if !cells.is_empty() {
            // Group cells by row using BTreeMap for sorted order
            let mut rows_map: BTreeMap<u32, BTreeMap<u32, &CellValue>> = BTreeMap::new();
            for ((row, col), value) in cells.iter() {
                rows_map
                    .entry(*row)
                    .or_default()
                    .insert(*col, value);
            }

            // Write rows in order
            for (row, cols) in rows_map {
                let mut row_elem = BytesStart::new("row");
                row_elem.push_attribute(("r", (row + 1).to_string().as_str()));
                writer.write_event(Event::Start(row_elem))?;

                for (col, value) in cols {
                    self.write_cell(&mut writer, row, col, value)?;
                }

                writer.write_event(Event::End(BytesEnd::new("row")))?;
            }
        }

        writer.write_event(Event::End(BytesEnd::new("sheetData")))?;
        writer.write_event(Event::End(BytesEnd::new("worksheet")))?;

        zip.write_all(writer.into_inner().into_inner().as_slice())?;
        Ok(())
    }

    fn write_cell(
        &self,
        writer: &mut Writer<Cursor<Vec<u8>>>,
        row: u32,
        col: u32,
        value: &CellValue,
    ) -> Result<()> {
        let cell_ref = cell_reference(row, col);

        match value {
            CellValue::Empty => {}
            CellValue::String(s) => {
                let mut cell = BytesStart::new("c");
                cell.push_attribute(("r", cell_ref.as_str()));
                cell.push_attribute(("t", "inlineStr"));
                writer.write_event(Event::Start(cell))?;

                writer.write_event(Event::Start(BytesStart::new("is")))?;
                writer.write_event(Event::Start(BytesStart::new("t")))?;
                writer.write_event(Event::Text(BytesText::new(s)))?;
                writer.write_event(Event::End(BytesEnd::new("t")))?;
                writer.write_event(Event::End(BytesEnd::new("is")))?;

                writer.write_event(Event::End(BytesEnd::new("c")))?;
            }
            CellValue::Number(n) => {
                let mut cell = BytesStart::new("c");
                cell.push_attribute(("r", cell_ref.as_str()));
                writer.write_event(Event::Start(cell))?;

                writer.write_event(Event::Start(BytesStart::new("v")))?;
                writer.write_event(Event::Text(BytesText::new(&n.to_string())))?;
                writer.write_event(Event::End(BytesEnd::new("v")))?;

                writer.write_event(Event::End(BytesEnd::new("c")))?;
            }
            CellValue::Boolean(b) => {
                let mut cell = BytesStart::new("c");
                cell.push_attribute(("r", cell_ref.as_str()));
                cell.push_attribute(("t", "b"));
                writer.write_event(Event::Start(cell))?;

                writer.write_event(Event::Start(BytesStart::new("v")))?;
                writer.write_event(Event::Text(BytesText::new(if *b { "1" } else { "0" })))?;
                writer.write_event(Event::End(BytesEnd::new("v")))?;

                writer.write_event(Event::End(BytesEnd::new("c")))?;
            }
        }

        Ok(())
    }
}
