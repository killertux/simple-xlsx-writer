use simple_xlsx_writer::{Row, WorkBook};
use std::fs::File;
use std::io::Write;

fn main() -> std::io::Result<()> {
    let mut files = File::create("tmp/example.xlsx")?;
    let mut workbook = WorkBook::new(&mut files)?;
    for s in 0..1 {
        println!("Sheet {}", s);
        let header_style = workbook.create_cell_style((255, 255, 255), (0, 0, 0));
        let mut sheet = workbook.get_new_sheet();
        sheet.write_sheet(|sheet_writer| {
            for x in 0..10 {
                if x % 5000 == 0 {
                    println!("Row {}", x);
                }
                let mut row = Row::new();
                for i in 0..50 {
                    if x == 0 {
                        row.add_cell((x as f64 + i as f64 / 100.0, &header_style).into());
                    } else {
                        row.add_cell((x as f64 + i as f64 / 100.0).into());
                    }
                }
                sheet_writer.write_row(row)?;
            }
            Ok(())
        })?;
    }
    workbook.finish()?;
    files.flush()?;
    Ok(())
}
