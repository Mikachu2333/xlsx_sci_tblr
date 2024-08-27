use calamine::{open_workbook, DataType, Reader, Xlsx};
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Workbook};
use std::path::PathBuf;

struct ExcelData {
    vec_data: Vec<Vec<String>>,
    row_index: usize,
    column_index: usize,
    header: usize,
    file_dir: PathBuf,
    file_name: String,
    file_ext: String,
}

fn main() {
    let args: Vec<String> = std::env::args().collect();

    if args.len() == 1 {
        println!(
            r#"
USAGE OF <xlsx_sci_tblr>

This program aims to assist users to foramt their tabulars scientifically.

Just drag your Excel table(s) to it, and the program would output it to where it from with a suffix name <_formatted>. 
"#
        );
        std::process::exit(0);
    } else if args.len() == 2 {
        let xlsx_path = PathBuf::from(&args[1]);
        if !xlsx_path.exists() {
            println!("NOT EXIST: <{}>", xlsx_path.display());
            return;
        }
        format_table(xlsx_path, 1);
    } else if args.len() == 3 {
        let xlsx_path = PathBuf::from(&args[1]);
        if !xlsx_path.exists() {
            panic!("NOT EXIST: <{}>", xlsx_path.display());
        }
        let header = match args[2].parse::<usize>() {
            Ok(x) => x,
            Err(_) => {
                panic!("HEADER ARG NOT NUM.");
            }
        };
        format_table(xlsx_path, header);
    } else {
        panic!("WRONG ARGS.");
    }
}

///Read excel data to `Vec<Vec<String>>` and packed infos to ExcelData struct
fn read_data(excel_path: PathBuf, header_row: usize) -> ExcelData {
    let mut packed_data = ExcelData {
        vec_data: Vec::new(),
        row_index: 0,
        column_index: 0,
        header: header_row - 1,
        file_dir: match excel_path.parent() {
            Some(p) => p.to_path_buf(),
            None => PathBuf::from(""),
        },
        file_name: match excel_path.file_stem() {
            Some(c) => c.to_string_lossy().to_string(),
            None => "".to_string(),
        },
        file_ext: match excel_path.extension() {
            Some(c) => c.to_string_lossy().to_string(),
            None => "".to_string(),
        },
    };

    packed_data.file_dir = match excel_path.parent() {
        Some(p) => p.to_path_buf(),
        None => PathBuf::from(""),
    };

    let mut datas: Vec<Vec<String>> = Vec::new();
    let exts = ["xlsx", "xls", "xlsm", "ods"];
    if !exts.contains(&packed_data.file_ext.as_str()) {
        panic!("ERROR FORMAT: <{}>", &packed_data.file_ext)
    };

    let mut excel: Xlsx<_> = match open_workbook(excel_path) {
        Ok(x) => x,
        Err(e) => panic!("ERROR OPEN FILE: <{}>", e),
    };

    if let Ok(range) = excel.worksheet_range(&excel.sheet_names()[0]) {
        packed_data.column_index = range.get_size().1 - 1;
        packed_data.row_index = range.get_size().0 - 1;
        //println!("{:?}", row_data);

        let mut temp_row_data: Vec<String> = Vec::new();
        for (_row_index, col_index, content) in range.cells() {
            match content.as_string() {
                Some(s) => temp_row_data.push(s),
                None => temp_row_data.push("".to_string()),
            };
            if (col_index + 1) % (packed_data.column_index + 1) == 0 {
                datas.push(temp_row_data);
                temp_row_data = Vec::new();
            }
        }
        packed_data.vec_data = datas
    }
    packed_data
}

fn remake_xlsx(excel_data: ExcelData) {
    let mut excel = Workbook::new();
    let sheet1 = excel.add_worksheet();

    let center = Format::new()
        .set_align(FormatAlign::Center)
        .set_align(FormatAlign::VerticalCenter);
    let top_line = center.clone().set_border_top(FormatBorder::Medium);
    let header_line = center.clone().set_border_bottom(FormatBorder::Thin);
    let top_header_line = center
        .clone()
        .set_border_top(FormatBorder::Medium)
        .set_border_bottom(FormatBorder::Thin);
    let bottom_line = center.clone().set_border_bottom(FormatBorder::Medium);

    for row in 0..=excel_data.row_index {
        if row == 0 && 0 == excel_data.header {
            match sheet1.write_row_with_format(
                row as u32,
                0,
                &excel_data.vec_data[row],
                &top_header_line,
            ) {
                Ok(_) => {}
                Err(e) => panic!("ERROR_WRITE_TO_TABLE{}", e),
            };
        } else if row == 0 && 0 != excel_data.header {
            match sheet1.write_row_with_format(row as u32, 0, &excel_data.vec_data[row], &top_line)
            {
                Ok(_) => {}
                Err(e) => panic!("ERROR_WRITE_TO_TABLE{}", e),
            };
        } else if row == excel_data.header {
            match sheet1.write_row_with_format(
                row as u32,
                0,
                &excel_data.vec_data[row],
                &header_line,
            ) {
                Ok(_) => {}
                Err(e) => panic!("ERROR_WRITE_TO_TABLE{}", e),
            };
        } else if row == excel_data.row_index {
            match sheet1.write_row_with_format(
                row as u32,
                0,
                &excel_data.vec_data[row],
                &bottom_line,
            ) {
                Ok(_) => {}
                Err(e) => panic!("ERROR_WRITE_TO_TABLE{}", e),
            };
        } else {
            match sheet1.write_row_with_format(row as u32, 0, &excel_data.vec_data[row], &center) {
                Ok(_) => {}
                Err(e) => panic!("ERROR_WRITE_TO_TABLE{}", e),
            };
        }
    }

    let new_name = excel_data.file_name + "_formatted." + &excel_data.file_ext;
    let new_path = excel_data.file_dir.join(new_name);
    match excel.save(new_path) {
        Ok(_) => {}
        Err(e) => panic!("ERROR WRITE OUTPUT: {}", e),
    };
}

///Format with given header
fn format_table(file_path: PathBuf, header: usize) {
    let raw_data = read_data(file_path, header);
    remake_xlsx(raw_data);
}
