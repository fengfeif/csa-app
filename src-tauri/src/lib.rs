use std::path::PathBuf;

#[tauri::command]
fn get_desktop_path() -> String {
    dirs::desktop_dir()
        .map(|p: PathBuf| p.to_string_lossy().to_string())
        .unwrap_or_else(|| String::from("/Users/fff/Desktop"))
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_http::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![get_desktop_path])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
