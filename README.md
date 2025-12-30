# outlook-cli

[![CI](https://github.com/Osso/outlook-cli/actions/workflows/ci.yml/badge.svg)](https://github.com/Osso/outlook-cli/actions/workflows/ci.yml)
[![GitHub release](https://img.shields.io/github/v/release/Osso/outlook-cli)](https://github.com/Osso/outlook-cli/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

CLI for Microsoft Graph Mail API access.

## Installation

```bash
cargo install --path .
```

## Setup

```bash
outlook login  # Opens browser for OAuth
```

## Usage

```bash
outlook list                    # List inbox messages
outlook list --unread           # List unread messages
outlook read <id>               # Read a specific message
outlook archive <id>            # Move to Archive folder
outlook spam <id>               # Move to Junk
outlook label <id> <category>   # Add category
outlook delete <id>             # Move to Deleted Items
outlook unsubscribe <id>        # Open unsubscribe link
```

## License

MIT
