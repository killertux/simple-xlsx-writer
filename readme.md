# simple_xlsx_writer
This is a very simple XLSX writer library.

This is not feature rich and it is not supposed to be. A lot of the design was based on the work of [simple_excel_writer](https://docs.rs/simple_excel_writer/latest/simple_excel_writer/) and I recomend you to check that crate.

The main idea of this create is to help you build XLSX files using very little RAM.
I created it to use in my web assembly site [csv2xlsx](https://csv2xlsx.com).

Basically, you just need to pass an output that implements [Write](std::io::Write) and [Sink](std::io::Sink) to the [WorkBook](crate::WorkBook). And while you are writing the file, it wil be written directly to the output already compressed. So, you could stream directly into a file using very little RAM. Or even write to the memory and still not use that much memory as the file will be already compressed.

## Example
```rust
use simple_xlsx_writer::{row, Row, WorkBook};
use std::fs::File;
use std::io::Write;

fn main() -> std::io::Result<()> {
 let mut files = File::create("example.xlsx")?;
 let mut workbook = WorkBook::new(&mut files)?;
 let header_style = workbook.create_cell_style((255, 255, 255), (0, 0, 0));
 workbook.get_new_sheet().write_sheet(|sheet_writer| {
     sheet_writer.write_row(row![("My", &header_style), ("Sample", &header_style), ("Header", &header_style)])?;
     sheet_writer.write_row(row![1, 2, 3])?;
     Ok(())
 })?;
 workbook.get_new_sheet().write_sheet(|sheet_writer| {
     sheet_writer.write_row(row![("Another", &header_style), ("Sheet", &header_style), ("Header", &header_style)])?;
     sheet_writer.write_row(row![1.32, 2.43, 3.54])?;
     Ok(())
 })?;
 workbook.finish()?;
 files.flush()?;
 Ok(())
}
```