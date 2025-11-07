use windows::{
	Win32::System::Diagnostics::Debug::{FORMAT_MESSAGE_FROM_SYSTEM, FORMAT_MESSAGE_IGNORE_INSERTS, FormatMessageW},
	core::{HRESULT, PWSTR},
};

#[derive(Debug)]
pub struct ComError {
	pub hr: HRESULT,
	pub ctx: String,
}

impl ComError {
	pub fn new(hr: HRESULT, ctx: impl Into<String>) -> Self {
		Self { hr, ctx: ctx.into() }
	}

	#[must_use]
	pub fn to_display(&self) -> String {
		let msg = format_message(self.hr);
		let code = u32::from_ne_bytes(self.hr.0.to_ne_bytes());
		if self.ctx.is_empty() {
			if msg.is_empty() { format!("HRESULT 0x{code:08X}") } else { format!("HRESULT 0x{code:08X} ({msg})") }
		} else if msg.is_empty() {
			format!("{}: HRESULT 0x{code:08X}", self.ctx)
		} else {
			format!("{}: HRESULT 0x{code:08X} ({msg})", self.ctx)
		}
	}
}

fn format_message(hr: HRESULT) -> String {
	const BUF_LEN_U32: u32 = 1024;
	let mut buf = [0u16; BUF_LEN_U32 as usize];
	unsafe {
		let len = FormatMessageW(
			FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
			None,
			u32::from_ne_bytes(hr.0.to_ne_bytes()),
			0,
			PWSTR(buf.as_mut_ptr()),
			BUF_LEN_U32,
			None,
		);
		if len == 0 {
			return String::new();
		}
		let s = String::from_utf16_lossy(&buf[..len as usize]);
		s.trim().trim_end_matches(['\r', '\n']).trim_end().to_string()
	}
}
