use embed_manifest::{
	embed_manifest,
	manifest::{ExecutionLevel::RequireAdministrator, SupportedOS::Windows10},
	new_manifest,
};

fn main() {
	let manifest =
		new_manifest("trypsynth.uphide").requested_execution_level(RequireAdministrator).supported_os(Windows10..);
	embed_manifest(manifest).expect("Failed to embed the application manifest");
	println!("cargo:rerun-if-changed=build.rs");
}
