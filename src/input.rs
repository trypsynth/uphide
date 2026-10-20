use std::ffi::OsString;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Mode {
	Hide,
	Restore,
}

impl Mode {
	pub const fn verb(self) -> &'static str {
		match self {
			Self::Hide => "Hide",
			Self::Restore => "Restore",
		}
	}

	pub const fn completed(self) -> &'static str {
		match self {
			Self::Hide => "Hidden",
			Self::Restore => "Restored",
		}
	}

	pub const fn criteria(self) -> &'static str {
		match self {
			Self::Hide => "IsInstalled=0 And IsHidden=0",
			Self::Restore => "IsInstalled=0 And IsHidden=1",
		}
	}
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Command {
	Menu,
	Run(Mode),
	Help,
	Version,
}

pub fn parse_command(arguments: impl IntoIterator<Item = OsString>) -> Result<Command, String> {
	let arguments: Vec<_> = arguments.into_iter().collect();
	match arguments.as_slice() {
		[] => Ok(Command::Menu),
		[argument] if argument == "--hide" => Ok(Command::Run(Mode::Hide)),
		[argument] if argument == "--restore" => Ok(Command::Run(Mode::Restore)),
		[argument] if argument == "--help" || argument == "-h" => Ok(Command::Help),
		[argument] if argument == "--version" || argument == "-V" => Ok(Command::Version),
		_ => Err("Use uphide, uphide --hide, or uphide --restore. Use --help for help.".to_owned()),
	}
}

pub fn parse_selection(input: &str, count: usize) -> Result<Vec<usize>, String> {
	let mut indices = Vec::new();
	for token in input.split_whitespace() {
		let number = token.parse::<usize>().map_err(|_| format!("'{token}' is not a valid update number."))?;
		if number == 0 || number > count {
			return Err(format!("'{token}' is outside the update list."));
		}
		indices.push(number - 1);
	}
	indices.sort_unstable();
	indices.dedup();
	Ok(indices)
}

#[cfg(test)]
mod tests {
	use std::ffi::OsString;

	use super::{Command, Mode, parse_command, parse_selection};

	#[test]
	fn modes_use_the_corresponding_hidden_state() {
		assert_eq!(Mode::Hide.criteria(), "IsInstalled=0 And IsHidden=0");
		assert_eq!(Mode::Restore.criteria(), "IsInstalled=0 And IsHidden=1");
		assert_eq!(Mode::Hide.verb(), "Hide");
		assert_eq!(Mode::Restore.verb(), "Restore");
		assert_eq!(Mode::Hide.completed(), "Hidden");
		assert_eq!(Mode::Restore.completed(), "Restored");
	}

	#[test]
	fn selection_is_sorted_unique_and_zero_based() {
		assert_eq!(parse_selection("3 1 3 2", 3), Ok(vec![0, 1, 2]));
	}

	#[test]
	fn invalid_selection_rejects_the_whole_request() {
		for input in ["1 nope", "0", "-1", "4", "99999999999999999999999999999999"] {
			assert!(parse_selection(input, 3).is_err(), "{input}");
		}
		assert!(parse_selection("1", 0).is_err());
	}

	#[test]
	fn empty_selection_cancels() {
		assert_eq!(parse_selection(" \t\n", 3), Ok(Vec::new()));
	}

	#[test]
	fn command_requires_one_known_argument() {
		assert_eq!(parse_command([]), Ok(Command::Menu));
		assert_eq!(parse_command([OsString::from("--restore")]), Ok(Command::Run(Mode::Restore)));
		assert_eq!(parse_command([OsString::from("--hide")]), Ok(Command::Run(Mode::Hide)));
		assert_eq!(parse_command([OsString::from("--help")]), Ok(Command::Help));
		assert!(parse_command([OsString::from("--hide"), OsString::from("--restore")]).is_err());
		assert!(parse_command([OsString::from("--unknown")]).is_err());
	}
}
