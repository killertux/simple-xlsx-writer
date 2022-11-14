mod row;
mod sheet;
mod workbook;

pub use row::{CellValue, Row};
pub use sheet::{Sheet, SheetWriter};
pub use workbook::{CellStyle, WorkBook};

#[macro_export]
macro_rules! row {
    ($( $x:expr ),*) => {
        {
            let mut row = Row::new();
            $(row.add_cell($x.into());)*
            row
        }
    };
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Result as IoResult};

    #[test]
    fn it_works() -> IoResult<()> {
        let mut cursor = Cursor::new(Vec::new());
        let mut workbook = WorkBook::new(&mut cursor)?;
        let cell_style = workbook.create_cell_style((255, 255, 255), (0, 0, 0));
        let mut sheet_1 = workbook.get_new_sheet();
        sheet_1.write_sheet(|sheet_writer| {
            sheet_writer.write_row(row!(
                (1, &cell_style),
                (10.3, &cell_style),
                (54.3, &cell_style)
            ))?;
            sheet_writer.write_row(row!("ola", "text", "tree"))?;
            sheet_writer.write_row(row!(true, false, false, false))?;
            Ok(())
        })?;
        let mut sheet_2 = workbook.get_new_sheet();
        sheet_2.write_sheet(|sheet_writer| {
            sheet_writer.write_row(row!(1, 2, 3, 4, 4))?;
            sheet_writer.write_row(row!("one", "two", "three"))?;
            sheet_writer.write_row(row!("Another row"))?;
            Ok(())
        })?;
        workbook.finish()?;
        Ok(())
    }
}
