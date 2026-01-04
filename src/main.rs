#![warn(clippy::all, clippy::pedantic, clippy::nursery)]

mod cli;
mod com;
mod error;
mod update;

use std::{
	io::{self, Write},
	process,
};

use error::ComError;
use update::UpdateManager;

fn main() {
	if let Err(e) = run() {
		eprintln!("COM Error: {}", e.to_display());
		eprintln!("This program requires Windows Update services to be running.");
		process::exit(1);
	}
}

fn run() -> Result<(), ComError> {
	let mut manager = UpdateManager::new()?;
	manager.search_for_updates()?;
	if !manager.has_updates() {
		print!("No updates available.\nPress Enter to exit...");
		let _ = Write::flush(&mut io::stdout());
		let _ = io::stdin().read_line(&mut String::new());
		return Ok(());
	}
	manager.display_available_updates();
	cli::print_instructions();
	let user_input = cli::get_user_input();
	if user_input.is_empty() {
		return Ok(());
	}
	let (mut indices, invalid) = cli::parse_user_input(&user_input);
	for tok in invalid {
		println!("Warning: '{tok}' is not a valid number. Skipping.");
	}
	if indices.is_empty() {
		return Ok(());
	}
	indices.sort_unstable();
	indices.dedup();
	println!("\nProceeding to hide selected updates...\n");
	manager.hide_selected_updates(&indices);

	cli::press_enter_to_exit();
	Ok(())
}
