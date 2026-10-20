use std::{marker::PhantomData, rc::Rc};

use windows::Win32::System::Com::{COINIT_APARTMENTTHREADED, CoInitializeEx, CoUninitialize};

use crate::error::ComError;

pub struct ComInit {
	_apartment: PhantomData<Rc<()>>,
}

impl ComInit {
	pub fn new() -> Result<Self, ComError> {
		unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }
			.ok()
			.map_err(|error| ComError::new(error, "Failed to initialize COM"))?;
		Ok(Self { _apartment: PhantomData })
	}
}

impl Drop for ComInit {
	fn drop(&mut self) {
		unsafe { CoUninitialize() }
	}
}
