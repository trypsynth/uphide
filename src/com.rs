use windows::{
	Win32::{
		Foundation::RPC_E_CHANGED_MODE,
		System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize},
	},
	core::HRESULT,
};

use crate::error::ComError;

pub struct ComInit {
	uninit: bool,
}

impl ComInit {
	/// Creates a new COM apartment initializer.
	///
	/// # Errors
	/// Returns a [`ComError`] if COM initialization fails in this thread.
	pub fn new() -> Result<Self, ComError> {
		unsafe {
			let hr: HRESULT = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
			if hr.is_ok() {
				Ok(Self { uninit: true })
			} else if hr == RPC_E_CHANGED_MODE {
				Ok(Self { uninit: false })
			} else {
				Err(ComError::new(hr, "Failed to initialize COM"))
			}
		}
	}
}

impl Drop for ComInit {
	fn drop(&mut self) {
		if self.uninit {
			unsafe { CoUninitialize() }
		}
	}
}
