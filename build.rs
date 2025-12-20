
fn main() {
    slint_build::compile("ui/main.slint").unwrap();

    #[cfg(target_os = "windows")]
    {
        let mut res = winres::WindowsResource::new();
        res.set_icon("assets/app.ico"); // path to your .ico
        res.compile().unwrap();
    }
}
