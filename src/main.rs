use csv::Reader;
use inquire::{InquireError, Select};
use std::error::Error;
use std::fs::{self, File};
use std::io::{BufRead, BufReader, Lines};
use std::iter::Peekable;
use std::path::PathBuf;
use std::time::Instant;
use xlsxwriter::*;

fn main() -> Result<(), Box<dyn Error>> {
    let options: Vec<&str> = vec!["Almadar", "Libyana", "LTT"];
    let ans: Result<&str, InquireError> = Select::new("What's the company?", options).prompt();
    let start = Instant::now();
    match ans {
        Ok(choice) => {
            if choice == "Libyana" {
                process_libyana(choice)?;
            } else if choice == "Almadar" {
                process_almadar(choice)?;
            } else {
                process_ltt(choice)?;
            }
        }
        Err(_) => println!("There was an error, please try again"),
    }
    println!("Time taken: {:?}", start.elapsed());

    Ok(())
}

fn process_libyana(dir_path: &str) -> Result<(), Box<dyn Error>> {
    println!("Value : Quantity");
    for entry in fs::read_dir(dir_path)? {
        let entry = entry?;
        let path = entry.path();
        if is_ext_file(&path, "dec") {
            process_libyana_file(&dir_path, &path)?;
        }
    }
    Ok(())
}
fn process_ltt(dir_path: &str) -> Result<(), Box<dyn Error>> {
    println!("Value : Quantity");
    for entry in fs::read_dir(dir_path)? {
        let entry = entry?;
        let path = entry.path();
        if is_ext_file(&path, "txt") {
            process_ltt_file(&dir_path, &path)?;
        }
    }
    Ok(())
}
fn process_ltt_file(dir_path: &str, path: &PathBuf) -> Result<(), Box<dyn Error>> {
    let file = File::open(path)?;
    let reader = BufReader::new(file);
    let mut rdr = reader.lines().peekable();
    while let Some(Ok(line)) = rdr.next() {
        if line.starts_with("FaceValue") {
            let face_value = line.split(":").collect::<Vec<&str>>()[1].trim();
            let value = face_value.parse::<i32>().unwrap() / 1000;
            let workbook_path = format!("{}/LTT {}.xlsx", dir_path, value);
            create_excel_ltt(&workbook_path, &mut rdr, value)?;
        }
    }

    Ok(())
}
fn create_excel_ltt(
    path: &str,
    rdr: &mut Peekable<Lines<BufReader<File>>>,
    value: i32,
) -> Result<(), Box<dyn Error>> {
    let workbook = Workbook::new(path)?;
    let mut sheet = workbook
        .add_worksheet(None)
        .map_err(|e| Box::new(e) as Box<dyn Error>)?;

    // Headers
    sheet.write_string(0, 0, "CARD_SEQ", None)?;
    sheet.write_string(0, 1, "CARD_SECRET", None)?;
    sheet.write_string(0, 2, "Value_Card", None)?;
    sheet.write_string(0, 3, "Code_Card", None)?;
    sheet.write_string(0, 4, "Exp_Card", None)?;
    sheet.write_string(0, 5, "Com_Name", None)?;

    let mut row: u32 = 1;
    let mut in_sequences = false;

    while let Some(Ok(line)) = rdr.next() {
        if line.trim() == "[BEGIN]" {
            in_sequences = true;
            continue;
        } else if line.trim() == "[END]" {
            break;
        }
        if in_sequences {
            if let Some((sequence, code)) = line.split_once(' ') {
                sheet.write_string(row, 0, &sequence.to_string(), None)?;
                sheet.write_string(row, 1, &code.to_string(), None)?;
                sheet.write_number(row, 2, value as f64, None)?;
                row += 1;
            }
        }
    }
    println!("{:?} : {:?}", value, row - 1,);
    workbook.close().map_err(|e| e.into())
}
fn process_libyana_file(dir_path: &str, path: &PathBuf) -> Result<(), Box<dyn Error>> {
    let file = File::open(path)?;
    let mut rdr = Reader::from_reader(file);
    if let Some(result) = rdr.records().next() {
        let record = result?;
        if let Some(value_str) = record.get(2) {
            if let Ok(num_str) = value_str.split('.').next().ok_or("Invalid input") {
                if let Ok(value) = num_str.parse::<i32>() {
                    let workbook_path = format!("{}/Libyana {}.xlsx", dir_path, value);
                    let mut pos = csv::Position::new();
                    pos.set_byte(0);
                    rdr.seek(pos)?;
                    create_excel_libyana(&workbook_path, &mut rdr, value)?;
                }
            }
        }
    }

    Ok(())
}

fn process_almadar_file(dir_path: &str, path: &PathBuf) -> Result<(), Box<dyn Error>> {
    let file = File::open(path)?;
    let mut rdr = Reader::from_reader(file);

    if let Some(result) = rdr.records().next() {
        let record = result?;
        if let Some(value_str) = record.get(3) {
            let value = value_str.parse::<i32>().unwrap() / 100;
            let workbook_path = format!("{}/Almadar {}.xlsx", dir_path, value);
            let mut pos = csv::Position::new();
            pos.set_byte(0);
            rdr.seek(pos)?;
            create_excel_almadar(&workbook_path, &mut rdr, value)?;
        }
    }

    Ok(())
}

fn create_excel_libyana(
    path: &str,
    rdr: &mut Reader<File>,
    value: i32,
) -> Result<(), Box<dyn Error>> {
    let workbook = Workbook::new(path)?;
    let mut sheet = workbook
        .add_worksheet(None)
        .map_err(|e| Box::new(e) as Box<dyn Error>)?;

    // Headers
    sheet.write_string(0, 0, "CARD_SEQ", None)?;
    sheet.write_string(0, 1, "CARD_SECRET", None)?;
    sheet.write_string(0, 2, "Value_Card", None)?;
    sheet.write_string(0, 3, "Code_Card", None)?;
    sheet.write_string(0, 4, "Exp_Card", None)?;
    sheet.write_string(0, 5, "Com_Name", None)?;

    let mut row: u32 = 1;
    for result in rdr.records() {
        let record = result?;
        sheet.write_string(row, 0, record.get(1).unwrap_or_default(), None)?;
        sheet.write_string(row, 1, record.get(0).unwrap_or_default(), None)?;
        sheet.write_number(row, 2, value as f64, None)?;
        row += 1;
    }
    println!("{:?} : {:?}", value, row - 1,);
    workbook.close().map_err(|e| e.into())
}

fn process_almadar(dir_path: &str) -> Result<(), Box<dyn Error>> {
    println!("Value : Quantity");
    for entry in fs::read_dir(dir_path)? {
        let entry = entry?;
        let path = entry.path();
        if is_ext_file(&path, "csv") {
            process_almadar_file(&dir_path, &path)?;
        }
    }
    Ok(())
}

fn create_excel_almadar(
    path: &str,
    rdr: &mut Reader<File>,
    value: i32,
) -> Result<(), Box<dyn Error>> {
    let workbook = Workbook::new(path)?;
    let mut sheet = workbook
        .add_worksheet(None)
        .map_err(|e| Box::new(e) as Box<dyn Error>)?;

    // Headers
    sheet.write_string(0, 0, "CARD_SEQ", None)?;
    sheet.write_string(0, 1, "CARD_SECRET", None)?;
    sheet.write_string(0, 2, "Value_Card", None)?;
    sheet.write_string(0, 3, "Code_Card", None)?;
    sheet.write_string(0, 4, "Exp_Card", None)?;
    sheet.write_string(0, 5, "Com_Name", None)?;

    let mut row: u32 = 1;
    for result in rdr.records() {
        let record = result?;
        sheet.write_string(row, 0, record.get(1).unwrap_or_default(), None)?;
        sheet.write_string(row, 1, record.get(0).unwrap_or_default(), None)?;
        sheet.write_number(row, 2, value as f64, None)?;
        row += 1;
    }
    println!("{:?} : {:?}", value, row - 1,);
    workbook.close().map_err(|e| e.into())
}

fn is_ext_file(path: &PathBuf, ext: &str) -> bool {
    path.extension().and_then(|s| s.to_str()) == Some(ext)
}
