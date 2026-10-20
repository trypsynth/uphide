use std::io::{IsTerminal, Result as IoResult, Write, stdin, stdout};

use crate::input::Mode;

pub fn print_help() {
	println!("Uphide {}", env!("CARGO_PKG_VERSION"));
	println!("Hide or restore individual Windows updates.");
	println!("\nUsage: uphide [--hide | --restore | --help | --version]");
	println!("\nWith no arguments, choose an action from the menu.");
	println!("Select updates by number, separated by spaces. Leave the selection blank to exit.");
	println!("Restoring an update makes it available again; it does not install it.");
}

pub fn read_input(prompt: &str) -> IoResult<String> {
	print!("{prompt}");
	stdout().flush()?;
	let mut input = String::new();
	stdin().read_line(&mut input)?;
	Ok(input.trim().to_owned())
}

pub fn choose_mode() -> IoResult<Option<Mode>> {
	println!("1. Hide updates\n2. Restore hidden updates");
	loop {
		match read_input("Choose 1 or 2, or press Enter to exit: ")?.as_str() {
			"1" => return Ok(Some(Mode::Hide)),
			"2" => return Ok(Some(Mode::Restore)),
			"" => return Ok(None),
			_ => println!("Enter 1 or 2."),
		}
	}
}

pub fn pause() -> IoResult<()> {
	if stdin().is_terminal() {
		read_input("\nPress Enter to exit...")?;
	}
	Ok(())
}
