// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use papyr_app::csv_parser;
use papyr_app::excel_parser;
use papyr_app::model::Document;
use papyr_app::parser::DocxParser;
use papyr_app::xlsx_model::{XlsxSheet, XlsxWorkbook};
use papyr_app::xlsx_parser::XlsxParser;
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;
use tauri::Manager;

const MAX_RECENT_FILES: usize = 8;
const SETTINGS_FILE_NAME: &str = "settings.json";
const RECENT_FILES_FILE_NAME: &str = "recent-files.json";

#[derive(Debug, Default)]
struct LaunchState {
    file_path: Mutex<Option<String>>,
}

#[derive(Debug, Serialize)]
#[serde(tag = "type", rename_all = "snake_case")]
enum OpenedFile {
    Docx { document: Document },
    Xlsx { workbook: XlsxWorkbook },
}

#[derive(Debug, Clone, Copy)]
enum FileKind {
    Docx,
    Spreadsheet(SpreadsheetKind),
}

#[derive(Debug, Clone, Copy)]
enum SpreadsheetKind {
    OpenXml,
    BinaryOrLegacy,
    Csv,
}

#[derive(Debug, Default, Serialize, Deserialize)]
struct AppSettings {
    theme: Option<String>,
}

#[derive(Debug, Default, Serialize, Deserialize)]
struct RecentFiles {
    paths: Vec<String>,
}

#[tauri::command]
fn open_docx(path: String, app: tauri::AppHandle) -> Result<Document, String> {
    let document = parse_docx_document(&path)?;
    remember_opened_file(&app, &path);
    Ok(document)
}

#[tauri::command]
fn open_file(path: String, app: tauri::AppHandle) -> Result<OpenedFile, String> {
    let file_kind = file_kind_from_path(Path::new(&path)).ok_or_else(|| {
        "Unsupported file type. Papyr can open .docx, .xlsx, .xlsm, .xlsb, .xls, and .csv files."
            .to_string()
    })?;

    let opened_file = match file_kind {
        FileKind::Docx => OpenedFile::Docx {
            document: parse_docx_document(&path)?,
        },
        FileKind::Spreadsheet(kind) => OpenedFile::Xlsx {
            workbook: parse_spreadsheet_workbook(&path, kind)?,
        },
    };

    remember_opened_file(&app, &path);
    Ok(opened_file)
}

#[tauri::command]
fn open_xlsx_sheet(path: String, sheet_index: usize) -> Result<XlsxSheet, String> {
    let Some(FileKind::Spreadsheet(kind)) = file_kind_from_path(Path::new(&path)) else {
        return Err("Unsupported file type. Expected an Excel workbook.".to_string());
    };

    parse_spreadsheet_sheet(&path, kind, sheet_index)
}

#[tauri::command]
fn open_in_microsoft_word(path: String) -> Result<(), String> {
    let document_path = PathBuf::from(path);

    if !is_docx_path(&document_path) {
        return Err("Only DOCX files can be opened in Microsoft Word.".to_string());
    }

    if !document_path.is_file() {
        return Err("The document no longer exists at its original location.".to_string());
    }

    launch_microsoft_word(&document_path)
}

fn is_docx_path(path: &Path) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case("docx"))
}

#[cfg(target_os = "windows")]
fn launch_microsoft_word(path: &Path) -> Result<(), String> {
    use std::os::windows::ffi::OsStrExt;
    use windows::core::{w, PCWSTR};
    use windows::Win32::Foundation::HWND;
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;

    let mut arguments = Vec::with_capacity(path.as_os_str().len() + 3);
    arguments.push('"' as u16);
    arguments.extend(path.as_os_str().encode_wide());
    arguments.push('"' as u16);
    arguments.push(0);

    let result = unsafe {
        ShellExecuteW(
            HWND::default(),
            w!("open"),
            w!("winword.exe"),
            PCWSTR(arguments.as_ptr()),
            PCWSTR::null(),
            SW_SHOWNORMAL,
        )
    };
    let result_code = result.0 as isize;

    if result_code > 32 {
        Ok(())
    } else if result_code == 2 {
        Err(
            "Microsoft Word could not be found. Make sure it is installed and try again."
                .to_string(),
        )
    } else {
        Err(format!(
            "Microsoft Word could not be opened (Windows error code {}).",
            result_code
        ))
    }
}

#[cfg(not(target_os = "windows"))]
fn launch_microsoft_word(_path: &Path) -> Result<(), String> {
    Err("Opening documents in Microsoft Word is currently supported on Windows only.".to_string())
}

fn parse_docx_document(path: &str) -> Result<Document, String> {
    println!("Opening DOCX file: {}", path);

    match DocxParser::from_path(path) {
        Ok(mut parser) => match parser.parse() {
            Ok(document) => {
                println!(
                    "Successfully parsed DOCX file: {} paragraphs, {} comments, {} images",
                    document.body.len(),
                    document.comments.len(),
                    document.images.len()
                );
                Ok(document)
            }
            Err(e) => {
                let error_msg = format!("Failed to parse DOCX file: {}", e);
                eprintln!("{}", error_msg);
                Err(error_msg)
            }
        },
        Err(e) => {
            let error_msg = format!("Failed to open DOCX file: {}", e);
            eprintln!("{}", error_msg);
            Err(error_msg)
        }
    }
}

fn parse_xlsx_workbook(path: &str) -> Result<XlsxWorkbook, String> {
    println!("Opening XLSX file: {}", path);

    match XlsxParser::from_path(path) {
        Ok(mut parser) => match parser.parse() {
            Ok(workbook) => {
                let active_rows = workbook
                    .active_sheet
                    .as_ref()
                    .map(|sheet| sheet.rows.len())
                    .unwrap_or_default();
                println!(
                    "Successfully parsed XLSX file: {} sheets, {} preview rows",
                    workbook.sheets.len(),
                    active_rows
                );
                Ok(workbook)
            }
            Err(e) => {
                let error_msg = format!("Failed to parse XLSX file: {}", e);
                eprintln!("{}", error_msg);
                Err(error_msg)
            }
        },
        Err(e) => {
            let error_msg = format!("Failed to open XLSX file: {}", e);
            eprintln!("{}", error_msg);
            Err(error_msg)
        }
    }
}

fn parse_spreadsheet_workbook(path: &str, kind: SpreadsheetKind) -> Result<XlsxWorkbook, String> {
    match kind {
        SpreadsheetKind::OpenXml => parse_xlsx_workbook(path),
        SpreadsheetKind::BinaryOrLegacy => {
            println!("Opening Excel file: {}", path);
            let workbook = excel_parser::parse_workbook(path)?;
            let active_rows = workbook
                .active_sheet
                .as_ref()
                .map(|sheet| sheet.rows.len())
                .unwrap_or_default();
            println!(
                "Successfully parsed Excel file: {} sheets, {} preview rows",
                workbook.sheets.len(),
                active_rows
            );
            Ok(workbook)
        }
        SpreadsheetKind::Csv => {
            println!("Opening CSV file: {}", path);
            let workbook = csv_parser::parse_workbook(path)?;
            let active_rows = workbook
                .active_sheet
                .as_ref()
                .map(|sheet| sheet.rows.len())
                .unwrap_or_default();
            println!(
                "Successfully parsed CSV file: {} preview rows",
                active_rows
            );
            Ok(workbook)
        }
    }
}

fn parse_spreadsheet_sheet(
    path: &str,
    kind: SpreadsheetKind,
    sheet_index: usize,
) -> Result<XlsxSheet, String> {
    match kind {
        SpreadsheetKind::OpenXml => {
            println!("Opening XLSX sheet {} from file: {}", sheet_index, path);
            let mut parser = XlsxParser::from_path(path)
                .map_err(|e| format!("Failed to open XLSX file: {}", e))?;
            parser
                .parse_sheet(sheet_index)
                .map_err(|e| format!("Failed to parse XLSX sheet: {}", e))
        }
        SpreadsheetKind::BinaryOrLegacy => {
            println!("Opening Excel sheet {} from file: {}", sheet_index, path);
            excel_parser::parse_sheet(path, sheet_index)
        }
        SpreadsheetKind::Csv => {
            println!("Opening CSV sheet {} from file: {}", sheet_index, path);
            csv_parser::parse_sheet(path, sheet_index)
        }
    }
}

#[tauri::command]
fn get_recent_files(app: tauri::AppHandle) -> Result<Vec<String>, String> {
    Ok(load_recent_files(&app)?)
}

#[tauri::command]
fn get_launch_file_path(state: tauri::State<LaunchState>) -> Result<Option<String>, String> {
    let mut path = state
        .file_path
        .lock()
        .map_err(|_| "Could not access launch state".to_string())?;
    Ok(path.take())
}

#[tauri::command]
fn get_theme_preference(app: tauri::AppHandle) -> Result<Option<String>, String> {
    Ok(load_settings(&app)?.theme)
}

#[tauri::command]
fn set_theme_preference(theme: String, app: tauri::AppHandle) -> Result<(), String> {
    let normalized_theme = normalize_theme(&theme)
        .ok_or_else(|| format!("Unsupported theme '{}'", theme))?
        .to_string();

    let mut settings = load_settings(&app)?;
    settings.theme = Some(normalized_theme);
    save_settings(&app, &settings)
}

#[tauri::command]
fn show_main_window(app: tauri::AppHandle) {
    if let Some(window) = app.get_webview_window("main") {
        let _ = window.show();
    }
}

#[tauri::command]
fn quit_app(app: tauri::AppHandle) {
    app.exit(0);
}

fn app_data_file_path(app: &tauri::AppHandle, filename: &str) -> Result<PathBuf, String> {
    let app_data_dir = app
        .path()
        .app_data_dir()
        .map_err(|e| format!("Could not resolve app data directory: {}", e))?;

    fs::create_dir_all(&app_data_dir)
        .map_err(|e| format!("Could not create app data directory: {}", e))?;

    Ok(app_data_dir.join(filename))
}

fn recent_files_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    app_data_file_path(app, RECENT_FILES_FILE_NAME)
}

fn settings_path(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    app_data_file_path(app, SETTINGS_FILE_NAME)
}

fn load_settings(app: &tauri::AppHandle) -> Result<AppSettings, String> {
    let path = settings_path(app)?;
    if !path.exists() {
        return Ok(AppSettings::default());
    }

    let content =
        fs::read_to_string(&path).map_err(|e| format!("Could not read settings: {}", e))?;

    serde_json::from_str(&content).map_err(|e| format!("Could not parse settings: {}", e))
}

fn save_settings(app: &tauri::AppHandle, settings: &AppSettings) -> Result<(), String> {
    let path = settings_path(app)?;
    let json = serde_json::to_string_pretty(settings)
        .map_err(|e| format!("Could not serialize settings: {}", e))?;

    fs::write(path, json).map_err(|e| format!("Could not write settings: {}", e))
}

fn load_recent_files(app: &tauri::AppHandle) -> Result<Vec<String>, String> {
    let path = recent_files_path(app)?;
    if !path.exists() {
        return Ok(Vec::new());
    }

    let content =
        fs::read_to_string(&path).map_err(|e| format!("Could not read recent files: {}", e))?;

    let mut recent_files: RecentFiles = serde_json::from_str(&content)
        .map_err(|e| format!("Could not parse recent files: {}", e))?;
    recent_files.paths.retain(|entry| !entry.trim().is_empty());
    Ok(recent_files.paths)
}

fn save_recent_files(app: &tauri::AppHandle, paths: &[String]) -> Result<(), String> {
    let path = recent_files_path(app)?;
    let payload = RecentFiles {
        paths: paths.to_vec(),
    };
    let json = serde_json::to_string_pretty(&payload)
        .map_err(|e| format!("Could not serialize recent files: {}", e))?;

    fs::write(path, json).map_err(|e| format!("Could not write recent files: {}", e))
}

fn remember_recent_file(app: &tauri::AppHandle, path: &str) -> Result<(), String> {
    let mut recent_files = load_recent_files(app).unwrap_or_default();
    recent_files.retain(|existing| existing != path);
    recent_files.insert(0, path.to_string());
    recent_files.truncate(MAX_RECENT_FILES);
    save_recent_files(app, &recent_files)
}

fn remember_opened_file(app: &tauri::AppHandle, path: &str) {
    if let Err(err) = remember_recent_file(app, path) {
        eprintln!("Failed to store recent file '{}': {}", path, err);
    }
}

fn normalize_theme(theme: &str) -> Option<&'static str> {
    match theme {
        "light" => Some("light"),
        "dark" => Some("dark"),
        _ => None,
    }
}

fn find_launch_file_path() -> Option<String> {
    std::env::args_os()
        .skip(1)
        .find_map(|arg| supported_file_path_from_arg(PathBuf::from(arg)))
}

fn supported_file_path_from_arg(path: PathBuf) -> Option<String> {
    if file_kind_from_path(&path).is_none() {
        return None;
    }

    let absolute_path = if path.is_absolute() {
        path
    } else {
        std::env::current_dir().ok()?.join(path)
    };

    if !absolute_path.exists() {
        return None;
    }

    Some(absolute_path.to_string_lossy().into_owned())
}

fn file_kind_from_path(path: &Path) -> Option<FileKind> {
    match path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase())
        .as_deref()
    {
        Some("docx") => Some(FileKind::Docx),
        Some("xlsx") | Some("xlsm") => Some(FileKind::Spreadsheet(SpreadsheetKind::OpenXml)),
        Some("xls") | Some("xlsb") => Some(FileKind::Spreadsheet(SpreadsheetKind::BinaryOrLegacy)),
        Some("csv") => Some(FileKind::Spreadsheet(SpreadsheetKind::Csv)),
        _ => None,
    }
}

fn main() {
    let launch_state = LaunchState {
        file_path: Mutex::new(find_launch_file_path()),
    };

    tauri::Builder::default()
        .manage(launch_state)
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![
            open_docx,
            open_file,
            open_in_microsoft_word,
            open_xlsx_sheet,
            get_recent_files,
            get_launch_file_path,
            get_theme_preference,
            set_theme_preference,
            show_main_window,
            quit_app
        ])
        .setup(|_app| Ok(()))
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
