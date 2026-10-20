mod cli;
mod com;
mod error;
mod input;
mod update;

use std::{env::args_os, error::Error, process::ExitCode};

use crate::{
	cli::{choose_mode, pause, print_help, read_input},
	input::{Command, parse_command, parse_selection},
	update::UpdateManager,
};

enum Outcome {
	Cancelled,
	Success,
	Failed,
}

fn main() -> ExitCode {
	let command = match parse_command(args_os().skip(1)) {
		Ok(Command::Help) => {
			print_help();
			return ExitCode::SUCCESS;
		}
		Ok(Command::Version) => {
			println!("Uphide {}", env!("CARGO_PKG_VERSION"));
			return ExitCode::SUCCESS;
		}
		Ok(command) => command,
		Err(error) => {
			eprintln!("{error}");
			return ExitCode::FAILURE;
		}
	};
	let (code, wait) = match run(command) {
		Ok(Outcome::Cancelled) => (ExitCode::SUCCESS, false),
		Ok(Outcome::Success) => (ExitCode::SUCCESS, true),
		Ok(Outcome::Failed) => (ExitCode::FAILURE, true),
		Err(error) => {
			eprintln!("Error: {error}");
			(ExitCode::FAILURE, true)
		}
	};
	if wait {
		if let Err(error) = pause() {
			eprintln!("Could not read input: {error}");
			return ExitCode::FAILURE;
		}
	}
	code
}

fn run(command: Command) -> Result<Outcome, Box<dyn Error>> {
	let mode = match command {
		Command::Run(mode) => mode,
		Command::Menu => match choose_mode()? {
			Some(mode) => mode,
			None => return Ok(Outcome::Cancelled),
		},
		Command::Help | Command::Version => return Ok(Outcome::Cancelled),
	};
	let manager = UpdateManager::new()?;
	let updates = manager.search(mode)?;
	if updates.is_empty() {
		println!("No matching updates found.");
		return Ok(Outcome::Success);
	}
	for (index, update) in updates.iter().enumerate() {
		println!("{}: {}", index + 1, update.title());
	}
	println!("\nEnter update numbers separated by spaces. Press Enter without a selection to exit.");
	let indices = loop {
		let input = read_input("Updates: ")?;
		match parse_selection(&input, updates.len()) {
			Ok(indices) if indices.is_empty() => return Ok(Outcome::Cancelled),
			Ok(indices) => break indices,
			Err(error) => println!("{error} Try again."),
		}
	};
	println!("\nSelected updates:");
	for &index in &indices {
		println!("{}", updates[index].title());
	}
	let answer = read_input(&format!("\n{} these updates? [y/N]: ", mode.verb()))?;
	if !answer.eq_ignore_ascii_case("y") && !answer.eq_ignore_ascii_case("yes") {
		return Ok(Outcome::Cancelled);
	}
	let mut failed = 0;
	for index in &indices {
		let update = &updates[*index];
		match update.apply(mode) {
			Ok(()) => println!("{}: {}", mode.completed(), update.title()),
			Err(error) => {
				failed += 1;
				eprintln!("Failed: {}: {error}", update.title());
			}
		}
	}
	println!("\n{} succeeded, {failed} failed.", indices.len() - failed);
	Ok(if failed == 0 { Outcome::Success } else { Outcome::Failed })
}
