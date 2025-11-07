use windows::{
	Win32::{
		Foundation::{E_POINTER, VARIANT_TRUE},
		System::{
			Com::{CLSCTX_INPROC_SERVER, CoCreateInstance},
			UpdateAgent::{ISearchResult, IUpdate, IUpdateCollection, IUpdateSearcher, IUpdateSession, UpdateSession},
		},
	},
	core::BSTR,
};

use crate::{com::ComInit, error::ComError};

pub struct UpdateInfo {
	pub title: String,
	pub index: i32,
	pub update: IUpdate,
}

pub struct UpdateManager {
	com: ComInit,
	session: IUpdateSession,
	searcher: IUpdateSearcher,
	search_result: Option<ISearchResult>,
	collection: Option<IUpdateCollection>,
	updates: Vec<UpdateInfo>,
}

impl UpdateManager {
	/// Creates a new update manager bound to a COM session.
	///
	/// # Errors
	/// Returns a [`ComError`] if creating the Windows Update COM objects fails.
	pub fn new() -> Result<Self, ComError> {
		let com = ComInit::new()?;
		unsafe {
			let session: IUpdateSession = CoCreateInstance(&UpdateSession, None, CLSCTX_INPROC_SERVER)
				.map_err(|e| ComError::new(e.code(), "Failed to create Update Session"))?;
			let searcher = session
				.CreateUpdateSearcher()
				.map_err(|e| ComError::new(e.code(), "Failed to create Update Searcher"))?;
			let id = BSTR::from("uphide");
			searcher
				.SetClientApplicationID(&id)
				.map_err(|e| ComError::new(e.code(), "Failed to set client application ID"))?;
			Ok(Self { com, session, searcher, search_result: None, collection: None, updates: Vec::new() })
		}
	}

	/// Searches for currently available, non-hidden updates.
	///
	/// # Errors
	/// Returns a [`ComError`] when the search fails through the underlying COM API.
	pub fn search_for_updates(&mut self) -> Result<(), ComError> {
		println!("Checking for updates...");
		unsafe {
			let criteria = BSTR::from("IsInstalled=0 And IsHidden=0");
			let result =
				self.searcher.Search(&criteria).map_err(|e| ComError::new(e.code(), "Failed to search for updates"))?;
			self.search_result = Some(result);
		}
		self.process_results()
	}

	fn process_results(&mut self) -> Result<(), ComError> {
		unsafe {
			let result = self.search_result.as_ref().ok_or_else(|| ComError::new(E_POINTER, "No search result"))?;
			let collection =
				result.Updates().map_err(|e| ComError::new(e.code(), "Failed to get update collection"))?;
			let count = collection.Count().map_err(|e| ComError::new(e.code(), "Failed to get update count"))?;
			self.collection = Some(collection.clone());
			self.updates.clear();
			for i in 0..count {
				let upd = collection.get_Item(i).map_err(|e| ComError::new(e.code(), "Failed to get update item"))?;
				let title = get_update_title(&upd);
				self.updates.push(UpdateInfo { title, index: i + 1, update: upd });
			}
		}
		Ok(())
	}

	#[must_use]
	pub const fn has_updates(&self) -> bool {
		!self.updates.is_empty()
	}

	pub fn display_available_updates(&self) {
		if self.updates.is_empty() {
			println!("No updates available.");
			return;
		}
		println!("Available updates:");
		for u in &self.updates {
			println!("{}: {}", u.index, u.title);
		}
	}

	pub fn hide_selected_updates(&self, indices: &[i32]) {
		if indices.is_empty() {
			return;
		}
		for &idx1 in indices {
			let array_index = idx1 - 1;
			let Ok(index) = usize::try_from(array_index) else {
				println!("Warning: Update index {idx1} is out of range. Skipping.");
				continue;
			};
			if index >= self.updates.len() {
				println!("Warning: Update index {idx1} is out of range. Skipping.");
				continue;
			}
			let info = &self.updates[index];
			unsafe {
				match info.update.SetIsHidden(VARIANT_TRUE) {
					Ok(()) => println!("Hid {}", info.title),
					Err(e) => {
						let hr = e.code();
						let code = u32::from_ne_bytes(hr.0.to_ne_bytes());
						println!("Failed to hide {}. (HRESULT: 0x{code:x})", info.title);
					}
				}
			}
		}
	}
}

fn get_update_title(update: &IUpdate) -> String {
	unsafe {
		update.Title().map_or_else(
			|_| "Unknown".to_string(),
			|b| {
				let s = b.to_string();
				if s.is_empty() { "Unknown".to_string() } else { s }
			},
		)
	}
}
