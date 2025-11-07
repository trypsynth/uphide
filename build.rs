use winresource::WindowsResource;

fn main() {
	let mut res = WindowsResource::new();
	res.set_manifest_file("uphide.manifest");
	res.compile().expect("failed to embed manifest");
}
