use windows::{
	Win32::{
		Foundation::{E_FAIL, VARIANT_FALSE, VARIANT_TRUE},
		System::{
			Com::{CLSCTX_INPROC_SERVER, CoCreateInstance},
			UpdateAgent::{
				IUpdate, IUpdateSearcher, IUpdateSession, UpdateSession, orcSucceeded, orcSucceededWithErrors,
			},
		},
	},
	core::{BSTR, Error as WindowsError},
};

use crate::{com::ComInit, error::ComError, input::Mode};

pub struct UpdateInfo {
	title: String,
	update: IUpdate,
}

impl UpdateInfo {
	pub fn title(&self) -> &str {
		&self.title
	}

	pub fn apply(&self, mode: Mode) -> Result<(), ComError> {
		let hidden = match mode {
			Mode::Hide => VARIANT_TRUE,
			Mode::Restore => VARIANT_FALSE,
		};
		unsafe { self.update.SetIsHidden(hidden) }
			.map_err(|error| ComError::new(error, "Windows could not change this update's hidden state"))
	}
}

pub struct UpdateManager {
	searcher: IUpdateSearcher,
	_session: IUpdateSession,
	// COM interfaces must be released before the apartment is uninitialized.
	_com: ComInit,
}

impl UpdateManager {
	pub fn new() -> Result<Self, ComError> {
		let com = ComInit::new()?;
		let session: IUpdateSession = unsafe { CoCreateInstance(&UpdateSession, None, CLSCTX_INPROC_SERVER) }
			.map_err(|error| ComError::new(error, "Failed to create the Windows Update session"))?;
		let searcher = unsafe { session.CreateUpdateSearcher() }
			.map_err(|error| ComError::new(error, "Failed to create the update searcher"))?;
		unsafe { searcher.SetClientApplicationID(&BSTR::from("Uphide")) }
			.map_err(|error| ComError::new(error, "Failed to set the client application ID"))?;
		Ok(Self { searcher, _session: session, _com: com })
	}

	pub fn search(&self, mode: Mode) -> Result<Vec<UpdateInfo>, ComError> {
		println!("Checking for updates...");
		unsafe {
			self.searcher
				.SetIncludePotentiallySupersededUpdates(if mode == Mode::Restore {
					VARIANT_TRUE
				} else {
					VARIANT_FALSE
				})
				.map_err(|error| ComError::new(error, "Failed to configure the update search"))?;
			let result = self
				.searcher
				.Search(&BSTR::from(mode.criteria()))
				.map_err(|error| ComError::new(error, "Failed to search for updates"))?;
			let code = result.ResultCode().map_err(|error| ComError::new(error, "Failed to read the search status"))?;
			if code == orcSucceededWithErrors {
				eprintln!("Windows returned a partial result. Some updates may be missing.");
			} else if code != orcSucceeded {
				return Err(ComError::new(
					WindowsError::from_hresult(E_FAIL),
					"Windows did not complete the update search",
				));
			}
			let collection = result.Updates().map_err(|error| ComError::new(error, "Failed to get the update list"))?;
			let count = collection.Count().map_err(|error| ComError::new(error, "Failed to get the update count"))?;
			let mut updates = Vec::new();
			for index in 0..count {
				let update =
					collection.get_Item(index).map_err(|error| ComError::new(error, "Failed to get an update"))?;
				let title =
					update.Title().map_err(|error| ComError::new(error, "Failed to get an update title"))?.to_string();
				updates.push(UpdateInfo { title, update });
			}
			Ok(updates)
		}
	}
}
