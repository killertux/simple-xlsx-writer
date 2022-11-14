use crate::Row;
use std::io::{Result as IoResult, Seek, Write};
use zip::{write::FileOptions, ZipWriter};

pub struct Sheet<'a, W>
where
    W: Write + Seek,
{
    id: usize,
    zip_writer: &'a mut ZipWriter<W>,
}

pub struct SheetWriter<W>
where
    W: Write,
{
    writer: W,
    row_index: usize,
    written_footer: bool,
}

impl<'a, W> Sheet<'a, W>
where
    W: Write + Seek,
{
    pub(crate) fn new(id: usize, zip_writer: &'a mut ZipWriter<W>) -> Self {
        Self { id, zip_writer }
    }

    pub fn write_sheet<T>(
        &mut self,
        function: impl FnOnce(&mut SheetWriter<&mut ZipWriter<W>>) -> IoResult<T>,
    ) -> IoResult<T> {
        let options = FileOptions::default();
        self.zip_writer
            .start_file(format!("xl/worksheets/sheet{}.xml", self.id), options)?;
        let mut sheet_writer = SheetWriter::start(&mut *self.zip_writer)?;
        let result = function(&mut sheet_writer)?;
        sheet_writer.finish()?;
        Ok(result)
    }

    pub fn sheet_writer(self) -> IoResult<SheetWriter<&'a mut ZipWriter<W>>> {
        let options = FileOptions::default();
        self.zip_writer
            .start_file(format!("xl/worksheets/sheet{}.xml", self.id), options)?;
        Ok(SheetWriter::start(&mut *self.zip_writer)?)
    }
}

impl<W> SheetWriter<W>
where
    W: Write,
{
    fn start(writer: W) -> IoResult<Self>
    where
        W: Write,
    {
        let mut writer = Self {
            writer,
            row_index: 0,
            written_footer: false,
        };
        writer.write_header()?;
        Ok(writer)
    }

    fn write_header(&mut self) -> IoResult<()> {
        write!(
            self.writer,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#
        )?;
        write!(self.writer, "\n<sheetData>\n")?;
        Ok(())
    }

    pub fn finish(mut self) -> IoResult<()> {
        self.write_footer()
    }

    fn write_footer(&mut self) -> IoResult<()> {
        self.written_footer = true;
        write!(self.writer, "\n</sheetData>\n</worksheet>\n")
    }

    pub fn write_row(&mut self, row: Row) -> IoResult<()> {
        self.row_index += 1;
        write!(self.writer, "<row r=\"{}\">\n", self.row_index)?;
        for (i, c) in row.cells().into_iter().enumerate() {
            c.write(i as u8, self.row_index, &mut self.writer)?;
        }
        write!(self.writer, "\n</row>\n")?;
        Ok(())
    }
}

impl<W> Drop for SheetWriter<W>
where
    W: Write,
{
    fn drop(&mut self) {
        if self.written_footer == false {
            self.write_footer().expect("Error written sheet footer");
        }
    }
}
