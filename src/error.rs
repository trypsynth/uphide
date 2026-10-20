use std::{
	error::Error,
	fmt::{Display, Formatter, Result as FmtResult},
};

use windows::core::Error as WindowsError;

#[derive(Debug)]
pub struct ComError {
	context: &'static str,
	source: WindowsError,
}

impl ComError {
	pub const fn new(source: WindowsError, context: &'static str) -> Self {
		Self { context, source }
	}
}

impl Display for ComError {
	fn fmt(&self, formatter: &mut Formatter<'_>) -> FmtResult {
		write!(formatter, "{}: {}", self.context, self.source)
	}
}

impl Error for ComError {
	fn source(&self) -> Option<&(dyn Error + 'static)> {
		Some(&self.source)
	}
}
