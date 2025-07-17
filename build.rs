fn main() {
    let mut res = winres::WindowsResource::new();
    res.set_icon("app.ico");
    res.set_resource_file("app.rc");
    res.compile().unwrap();
}
