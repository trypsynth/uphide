use std::io::{self, Write};

pub fn print_instructions() {
	println!("Enter the numbers of the updates you want to hide, separated by spaces.");
	println!("Press Enter without typing anything to exit.");
}

#[must_use]
pub fn get_user_input() -> String {
	let mut input = String::new();
	let _ = io::stdin().read_line(&mut input);
	input.trim().to_string()
}

#[must_use]
pub fn parse_user_input(input: &str) -> (Vec<i32>, Vec<String>) {
	let mut indices = Vec::new();
	let mut invalid = Vec::new();
	for tok in input.split_whitespace() {
		match tok.parse::<i32>() {
			Ok(n) if n > 0 => indices.push(n),
			_ => invalid.push(tok.to_string()),
		}
	}
	(indices, invalid)
}

pub fn press_enter_to_exit() {
	print!("\nPress Enter to exit...");
	let _ = io::stdout().flush();
	let _ = io::stdin().read_line(&mut String::new());
}
