//! Embeds the icon in the executable, for Explorer, the taskbar and the title bar.

fn main() {
    println!("cargo:rerun-if-changed=assets/mmv.rc");
    println!("cargo:rerun-if-changed=assets/mmv.ico");
    embed_resource::compile("assets/mmv.rc", embed_resource::NONE)
        .manifest_optional()
        .expect("compile the icon resource");
}
