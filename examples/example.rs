use simple_xlsx_writer::{row, Row, WorkBook};
use std::fs::File;
use std::io::Write;

fn main() -> std::io::Result<()> {
    let mut files = File::create("example.xlsx")?;
    let mut workbook = WorkBook::new(&mut files)?;
    let header_style = workbook.create_cell_style((255, 255, 255), (0, 0, 0));
    workbook.get_new_sheet().write_sheet(|sheet_writer| {
        sheet_writer.write_row(row![
            ("My", &header_style),
            ("Sample", &header_style),
            ("Header", &header_style)
        ])?;
        sheet_writer.write_row(row![1, 2, 3])?;
        Ok(())
    })?;
    workbook.get_new_sheet().write_sheet(|sheet_writer| {
        sheet_writer.write_row(row![
            ("Another", &header_style),
            ("Sheet", &header_style),
            ("Header", &header_style)
        ])?;
        sheet_writer.write_row(row![1.32, 2.43, 3.54])?;
        Ok(())
    })?;
    workbook.finish()?;
    files.flush()?;
    Ok(())
}
