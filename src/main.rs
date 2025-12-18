// src/main.rs
// BE-Alert xlsx to cvs convertor 
// -------------------------------
//
// File format tested with alken.be xlsx files
//
// Tools4Video BV All Rights reserved
// Copyright (c) 2025 Tools4Video 
// Developer Marc Colemont
//
//  XLSX, save CSV (; separated)
//
// Reads XLSX columns (required):
// - Voornaam
// - Naam
// - Straat
// - Huisnummer
// - Mobiel nummer
// - E-mailadres
//
// Outputs BE-Alert BIN NEW CSV format (33 columns):
// - 1st column "Tel/Ref." = formatted phone from "Mobiel nummer"
// - 2nd column "Civilité"
// - "Adres incl huisnummer" uses ONLY numeric part of Huisnummer (e.g. 11A -> 11, "12 Bus 3" -> 12)
// - NOTE: Output columns "Voornaam" and "Naam" must be swapped (provider error)
//
// Fixed values:
// - Postcode = BE-3570
// - Gemeente = Alken
// - Taal = NL
// - Land = BE
// - Type Contact = P



use anyhow::{anyhow, Result};
use calamine::{open_workbook, Data, Reader, Xlsx};
use csv::WriterBuilder;
use rfd::FileDialog;
use std::collections::HashMap;
use std::path::Path;
use slint::CloseRequestResponse;

const REQUIRED_COLUMNS: [&str; 6] = [
    "Voornaam",
    "Naam",
    "Straat",
    "Huisnummer",
    "Mobiel nummer",
    "E-mailadres",
];

slint::include_modules!();

fn cell_to_string(cell: &Data) -> String {
    match cell {
        Data::String(s) => s.clone(),
        Data::Float(f) => {
            if f.fract() == 0.0 {
                (*f as i64).to_string()
            } else {
                f.to_string()
            }
        }
        Data::Int(i) => i.to_string(),
        Data::Bool(b) => b.to_string(),
        Data::Empty => String::new(),
        _ => String::new(),
    }
}

fn get(cols: &HashMap<String, usize>, row: &[Data], name: &str) -> String {
    cols.get(name)
        .and_then(|&i| row.get(i))
        .map(cell_to_string)
        .unwrap_or_default()
        .trim()
        .to_string()
}

/// Keep only leading digits; stop at first non-digit.
/// Examples:
/// - "11A" -> "11"
/// - "12 Bus 3" -> "12"
fn extract_house_number(input: &str) -> String {
    let mut digits = String::new();
    for c in input.trim().chars() {
        if c.is_ascii_digit() {
            digits.push(c);
        } else {
            break;
        }
    }
    digits
}

/// Belgium-style normalization:
/// - "+32..." -> "0032..."
/// - "0..."   -> "0032..." (drop leading 0)
/// - strips spaces/dashes/etc (keeps digits and leading '+')
fn normalize_be_phone(input: &str) -> String {
    let mut s: String = input
        .trim()
        .chars()
        .filter(|c| c.is_ascii_digit() || *c == '+')
        .collect();

    if s.is_empty() {
        return String::new();
    }

    if s.starts_with("+32") {
        return format!("0032{}", &s[3..]);
    }

    if s.starts_with('0') {
        s.remove(0);
        return format!("0032{}", s);
    }

    if s.starts_with('4') {
        return format!("0032{}", s);
    }

    s
}

fn validate_xlsx_columns(input_xlsx: &str) -> Result<()> {
    let mut workbook: Xlsx<_> = open_workbook(input_xlsx)?;
    let range = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| anyhow!("No sheet found in XLSX"))??;

    let mut rows = range.rows();
    let header = rows
        .next()
        .ok_or_else(|| anyhow!("Empty sheet (no header row)"))?;

    let mut cols: HashMap<String, usize> = HashMap::new();
    for (i, cell) in header.iter().enumerate() {
        let name = cell_to_string(cell).trim().to_string();
        if !name.is_empty() {
            cols.insert(name, i);
        }
    }

    for required in REQUIRED_COLUMNS {
        if !cols.contains_key(required) {
            return Err(anyhow!("Missing required XLSX column: {}", required));
        }
    }

    Ok(())
}

fn convert_xlsx_to_csv(input_xlsx: &str, output_csv: &str) -> Result<()> {
    let mut workbook: Xlsx<_> = open_workbook(input_xlsx)?;
    let range = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| anyhow!("No sheet found in XLSX"))??;

    let mut rows = range.rows();

    // Header row -> column name -> index
    let header = rows.next().ok_or_else(|| anyhow!("Empty sheet (no header row)"))?;
    let mut cols: HashMap<String, usize> = HashMap::new();
    for (i, cell) in header.iter().enumerate() {
        let name = cell_to_string(cell).trim().to_string();
        if !name.is_empty() {
            cols.insert(name, i);
        }
    }

    // Required input columns
    for required in REQUIRED_COLUMNS {
        if !cols.contains_key(required) {
            return Err(anyhow!("Missing required XLSX column: {}", required));
        }
    }

    let mut writer = WriterBuilder::new()
        .delimiter(b';')
        .from_path(output_csv)?;

    // NEW output CSV header (33 columns)
    writer.write_record(&[
        "Tel/Ref.",
        "Civilité",
        "Naam",
        "Voornaam",
        "Adres incl huisnummer",
        "Bijkomend adres",
        "Postcode",
        "Gemeente",
        "Geboortedatum",
        "Email",
        "FAX",
        "FAX2",
        "FAX3",
        "Verdieping",
        "Aantal inwoners",
        "Telefoon 2",
        "Telefoon 3",
        "Telefoon 4",
        "Telefoon 5",
        "Telefoone 6",
        "Telefoon 7",
        "SMS",
        "SMS 2",
        "SMS 3",
        "Pager",
        "Zone libre 1",
        "Zone libre 2",
        "Zone libre 3",
        "Taal",
        "Land",
        "Rode lijst",
        "Type Contact",
        "GPS coördinaten",
    ])?;

    // Fixed values
    let fixed_postcode = "3570";
    let fixed_gemeente = "Alken";
    let fixed_taal = "NL";
    let fixed_land = "BE";
    let fixed_zwarte_lijst = "0";
    let fixed_type_contact = "P";

    for row in rows {
        // Read XLSX fields
        let xlsx_voornaam = get(&cols, row, "Voornaam");
        let xlsx_naam = get(&cols, row, "Naam");

        let straat = get(&cols, row, "Straat");
        let huisnr_raw = get(&cols, row, "Huisnummer");
        let huisnr_clean = extract_house_number(&huisnr_raw);

        let email = get(&cols, row, "E-mailadres");

        let mobiel_raw = get(&cols, row, "Mobiel nummer");
        let tel_ref = normalize_be_phone(&mobiel_raw);

        let adres_incl = format!("{} {}", straat, huisnr_clean).trim().to_string();

        // IMPORTANT: swap output fields (provider error)
        // CSV "Voornaam" <- XLSX "Naam"
        // CSV "Naam"     <- XLSX "Voornaam"
        let csv_voornaam = xlsx_voornaam;
        let csv_naam = xlsx_naam;

        writer.write_record([
            tel_ref,                       // Tel/Ref.
            String::new(),                 // Civilité
            csv_naam,                  // Naam 
            csv_voornaam,                   // VoorNaam  
            adres_incl,                    // Adres incl huisnummer
            String::new(),                 // Bijkomend adres
            fixed_postcode.to_string(),    // Postcode
            fixed_gemeente.to_string(),    // Gemeente
            String::new(),                 // Geboortedatum
            email,                         // Email
            String::new(),                 // FAX
            String::new(),                 // FAX2
            String::new(),                 // FAX3
            String::new(),                 // Verdieping
            String::new(),                 // Aantal inwoners
            String::new(),                 // Telefoon 2
            String::new(),                 // Telefoon 3
            String::new(),                 // Telefoon 4
            String::new(),                 // Telefoon 5
            String::new(),                 // Telefoon 6
            String::new(),                 // Telefoon 7
            String::new(),                 // SMS
            String::new(),                 // SMS 2
            String::new(),                 // SMS 3
            String::new(),                 // Pager
            String::new(),                 // Zone libre 1
            String::new(),                 // Zone libre 2
            String::new(),                 // Zone libre 3
            fixed_taal.to_string(),        // Taal
            fixed_land.to_string(),        // Land
            fixed_zwarte_lijst.to_string(),// Zwarte lijst
            fixed_type_contact.to_string(), // Type Contact
            String::new(),                 // GPS coördinaten
        ])?;
    }

    writer.flush()?;
    Ok(())
}

fn main() -> Result<()> {
    let ui = MainWindow::new()?;

    {
        let window = ui.window();
        window.on_close_requested(|| {
            let _ = slint::quit_event_loop();
            CloseRequestResponse::HideWindow
        });
    }

    ui.on_import_clicked({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                if let Some(file) = FileDialog::new()
                    .add_filter("Excel", &["xlsx"])
                    .pick_file()
                {
                    let path_str = file.display().to_string();
                    ui.set_input_file(path_str.clone().into());
                    ui.set_output_file("".into());
                    ui.set_export_checked(false);
                    ui.set_export_ok(false);

                    ui.set_import_checked(true);
                    match validate_xlsx_columns(&path_str) {
                        Ok(_) => {
                            ui.set_import_ok(true);
                            ui.set_status("XLSX selected and columns OK.".into());
                        }
                        Err(e) => {
                            ui.set_import_ok(false);
                            ui.set_status(format!("XLSX error: {}", e).into());
                        }
                    }
                }
            }
        }
    });

    ui.on_export_clicked({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                let input = ui.get_input_file().to_string();
                if input.trim().is_empty() {
                    ui.set_status("No XLSX selected.".into());
                    return;
                }

                let suggested_name = Path::new(&input)
                    .file_stem()
                    .and_then(|s| s.to_str())
                    .map(|stem| format!("{}.csv", stem))
                    .unwrap_or_else(|| "output.csv".to_string());

                if let Some(out) = FileDialog::new()
                    .add_filter("CSV", &["csv"])
                    .set_file_name(suggested_name)
                    .save_file()
                {
                    match convert_xlsx_to_csv(&input, out.to_str().unwrap()) {
                        Ok(_) => {
                            ui.set_output_file(out.display().to_string().into());
                            ui.set_status("CSV saved.".into());
                            ui.set_export_checked(true);
                            ui.set_export_ok(true);
                        }
                        Err(e) => {
                            ui.set_status(format!("Error: {}", e).into());
                            ui.set_export_checked(true);
                            ui.set_export_ok(false);
                        }
                    }
                }
            }
        }
    });

    ui.on_reset_clicked({
        let ui_handle = ui.as_weak();
        move || {
            if let Some(ui) = ui_handle.upgrade() {
                ui.set_input_file("".into());
                ui.set_output_file("".into());
                ui.set_status("Ready.".into());
                ui.set_import_checked(false);
                ui.set_import_ok(false);
                ui.set_export_checked(false);
                ui.set_export_ok(false);
            }
        }
    });

    ui.run()?;
    Ok(())
}
