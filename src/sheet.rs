use crate::Row;
use std::io::{Result as IoResult, Seek, Write};
use zip::{write::FileOptions, ZipWriter};

/// A XLSX sheet.
pub struct Sheet<'a, W>
where
    W: Write + Seek,
{
    id: usize,
    zip_writer: &'a mut ZipWriter<W>,
}

/// Responsible to write a sheet into the workbook.
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

    /// Receives a closure that will write the sheet. The closure receive a [SheetWriter](SheetWriter) that can be used to write the rows into the sheet.
    /// You don't need to call [finish](SheetWriter::finish) as it will be called for you.
    pub fn write_sheet<T>(
        self,
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

    /// Antoher way to write a sheet. Insted of using a closure that has access to the [SheetWriter](SheetWriter). This returns the [SheetWriter](SheetWriter) directly and you can use it to write the sheet.
    /// You need to call [finish](SheetWriter::finish).
    pub fn sheet_writer(self) -> IoResult<SheetWriter<&'a mut ZipWriter<W>>> {
        let options = FileOptions::default();
        self.zip_writer
            .start_file(format!("xl/worksheets/sheet{}.xml", self.id), options)?;
        SheetWriter::start(&mut *self.zip_writer)
    }
}

impl<W> SheetWriter<W>
where
    W: Write,
{
    /// Writes a row into the sheet.
    pub fn write_row(&mut self, row: Row) -> IoResult<()> {
        self.row_index += 1;
        writeln!(self.writer, "<row r=\"{}\">", self.row_index)?;
        for (i, c) in row.cells().into_iter().enumerate() {
            c.write(i as u8, self.row_index, &mut self.writer)?;
        }
        writeln!(self.writer, "\n</row>")?;
        Ok(())
    }

    /// Finish the sheet. Necessary to be called if you got the [SheetWriter](SheetWriter) from [Sheet::sheet_writer](Sheet::sheet_writer). We also try to execute this in the [Drop](SheetWriter::drop), but it is a good practice to always finish the sheet.
    pub fn finish(mut self) -> IoResult<()> {
        self.write_footer()
    }

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

    fn write_footer(&mut self) -> IoResult<()> {
        self.written_footer = true;
        write!(self.writer, "\n</sheetData>\n</worksheet>\n")
    }
}

impl<W> Drop for SheetWriter<W>
where
    W: Write,
{
    /// Drops the [SheetWriter](SheetWriter) and tries to finish it if not already finished. This might panic if we fail to write the footer of the sheet.
    fn drop(&mut self) {
        if !self.written_footer {
            self.write_footer().expect("Error written sheet footer");
        }
    }
}
