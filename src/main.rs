mod api;
mod auth;
mod config;

use anyhow::Result;
use clap::{Parser, Subcommand};

#[derive(Parser)]
#[command(name = "outlook")]
#[command(about = "CLI tool to access Microsoft Graph Mail API")]
struct Cli {
    /// Output as JSON
    #[arg(long, global = true)]
    json: bool,

    #[command(subcommand)]
    command: Commands,
}

#[derive(Subcommand)]
enum Commands {
    /// Set OAuth client credentials (from Azure App Registration)
    Config {
        /// Client ID (Application ID)
        client_id: String,
        /// Client Secret (optional, will prompt if not provided)
        client_secret: Option<String>,
    },
    /// Authenticate with Microsoft (opens browser)
    Login,
    /// List categories (like Gmail labels)
    Labels,
    /// List messages
    List {
        /// Maximum number of messages to show
        #[arg(short = 'n', long, default_value = "100")]
        max: u32,
        /// Search query
        #[arg(short, long)]
        query: Option<String>,
        /// Folder to filter by (inbox, sent, drafts, archive, trash, spam)
        #[arg(short, long, default_value = "inbox")]
        label: String,
        /// Show only unread messages
        #[arg(short, long)]
        unread: bool,
    },
    /// Read a specific message
    Read {
        /// Message ID
        id: String,
    },
    /// Archive a message (move to Archive folder)
    Archive {
        /// Message ID
        id: String,
    },
    /// Mark a message as spam (move to Junk)
    Spam {
        /// Message ID
        id: String,
    },
    /// Remove from spam and move to inbox
    Unspam {
        /// Message ID
        id: String,
    },
    /// Add a category to a message
    Label {
        /// Message ID
        id: String,
        /// Category to add
        label: String,
    },
    /// Remove a category from a message
    Unlabel {
        /// Message ID
        id: String,
        /// Category to remove
        label: String,
    },
    /// Move a message to trash (Deleted Items)
    Delete {
        /// Message ID
        id: String,
    },
    /// Unsubscribe from a mailing list (opens unsubscribe link)
    Unsubscribe {
        /// Message ID
        id: String,
    },
}

/// Normalize folder names to well-known folder names
fn normalize_folder(folder: &str) -> String {
    match folder.to_lowercase().as_str() {
        "inbox" => "inbox".to_string(),
        "sent" | "sentitems" => "sentitems".to_string(),
        "drafts" | "draft" => "drafts".to_string(),
        "trash" | "deleted" | "deleteditems" => "deleteditems".to_string(),
        "spam" | "junk" | "junkemail" => "junkemail".to_string(),
        "archive" => "archive".to_string(),
        "outbox" => "outbox".to_string(),
        other => other.to_string(),
    }
}

async fn get_client() -> Result<api::Client> {
    let cfg = config::load_config()?;
    let client_id = cfg.client_id.ok_or_else(|| {
        anyhow::anyhow!("Not configured. Run 'outlook config <client-id>' first")
    })?;
    let client_secret = cfg.client_secret.ok_or_else(|| {
        anyhow::anyhow!("Not configured. Run 'outlook config <client-id>' first")
    })?;

    let tokens = match config::load_tokens() {
        Ok(t) => t,
        Err(_) => anyhow::bail!("Not logged in. Run 'outlook login' first"),
    };

    // Try to use existing token, refresh if needed
    let client = api::Client::new(&tokens.access_token);

    // Test if token works by making a simple request
    match client.list_folders().await {
        Ok(_) => Ok(client),
        Err(_) => {
            // Token expired, try refresh
            let new_tokens = auth::refresh_token(&client_id, &client_secret, &tokens.refresh_token).await?;
            Ok(api::Client::new(&new_tokens.access_token))
        }
    }
}

#[tokio::main]
async fn main() -> Result<()> {
    let cli = Cli::parse();

    match cli.command {
        Commands::Config { client_id, client_secret } => {
            let secret = match client_secret {
                Some(s) => s,
                None => {
                    let s = rpassword::prompt_password("Client Secret: ")?;
                    if s.is_empty() {
                        anyhow::bail!("Client secret cannot be empty");
                    }
                    s
                }
            };

            let cfg = config::Config {
                client_id: Some(client_id),
                client_secret: Some(secret),
            };
            config::save_config(&cfg)?;
            println!("Credentials saved to {:?}", config::config_dir());
        }
        Commands::Login => {
            let cfg = config::load_config()?;
            let client_id = cfg.client_id.ok_or_else(|| {
                anyhow::anyhow!("Not configured. Run 'outlook config <client-id>' first")
            })?;
            let client_secret = cfg.client_secret.ok_or_else(|| {
                anyhow::anyhow!("Not configured. Run 'outlook config <client-id>' first")
            })?;

            // Delete existing tokens to force fresh login with new scopes
            let _ = std::fs::remove_file(config::tokens_path());

            auth::login(&client_id, &client_secret).await?;
            println!("Login successful! Tokens saved.");
        }
        Commands::Labels => {
            let client = get_client().await?;
            let categories = client.list_categories().await?;

            if let Some(cats) = categories.value {
                if cli.json {
                    println!("{}", serde_json::to_string(&cats)?);
                } else {
                    println!("Categories:");
                    for cat in cats {
                        let color = cat.color.as_deref().unwrap_or("none");
                        println!("  {} (color: {})", cat.display_name, color);
                    }
                }
            } else if cli.json {
                println!("[]");
            } else {
                println!("No categories found.");
            }
        }
        Commands::List { max, query, label, unread } => {
            let client = get_client().await?;
            let folder = normalize_folder(&label);

            let filter = if unread {
                Some("isRead eq false")
            } else {
                None
            };

            let list = if let Some(q) = &query {
                // Use search endpoint for query
                client.search_messages(q, max).await?
            } else {
                client.list_messages(&folder, filter, max).await?
            };

            if let Some(messages) = list.value {
                if cli.json {
                    let items: Vec<_> = messages.iter().map(|msg| {
                        serde_json::json!({
                            "id": msg.id,
                            "from": msg.get_from(),
                            "subject": msg.subject,
                            "date": msg.received_date_time,
                            "snippet": msg.body_preview,
                            "isRead": msg.is_read,
                            "categories": msg.categories,
                        })
                    }).collect();
                    println!("{}", serde_json::to_string(&items)?);
                } else {
                    for msg in messages {
                        let from = msg.get_from().unwrap_or_else(|| "Unknown".to_string());
                        let subject = msg.subject.as_deref().unwrap_or("(no subject)");
                        println!("{} | {} | {}", msg.id, from, subject);
                    }
                }
            } else if !cli.json {
                println!("No messages found.");
            } else {
                println!("[]");
            }
        }
        Commands::Read { id } => {
            let client = get_client().await?;
            let msg = client.get_message(&id).await?;

            if cli.json {
                println!("{}", serde_json::to_string(&serde_json::json!({
                    "id": msg.id,
                    "from": msg.get_from(),
                    "to": msg.get_to(),
                    "subject": msg.subject,
                    "date": msg.received_date_time,
                    "body": msg.get_body_text(),
                    "snippet": msg.body_preview,
                    "isRead": msg.is_read,
                    "categories": msg.categories,
                }))?);
            } else {
                println!("From: {}", msg.get_from().unwrap_or_else(|| "Unknown".to_string()));
                println!("To: {}", msg.get_to().unwrap_or_else(|| "Unknown".to_string()));
                println!("Subject: {}", msg.subject.as_deref().unwrap_or("(no subject)"));
                println!("Date: {}", msg.received_date_time.as_deref().unwrap_or("Unknown"));
                println!("---");

                if let Some(body) = msg.get_body_text() {
                    println!("{}", body);
                } else if let Some(preview) = &msg.body_preview {
                    println!("{}", preview);
                }
            }
        }
        Commands::Archive { id } => {
            let client = get_client().await?;
            client.archive(&id).await?;
            println!("Archived {}", id);
        }
        Commands::Spam { id } => {
            let client = get_client().await?;
            // Try to unsubscribe first, ignore errors (not all messages have unsubscribe)
            let msg = client.get_message(&id).await?;
            if let Some(url) = msg.get_unsubscribe_url() {
                if url.starts_with("http") {
                    let _ = open::that(&url);
                }
            }
            client.mark_spam(&id).await?;
            println!("Marked as spam {}", id);
        }
        Commands::Unspam { id } => {
            let client = get_client().await?;
            client.unspam(&id).await?;
            println!("Moved to inbox {}", id);
        }
        Commands::Label { id, label } => {
            let client = get_client().await?;
            client.add_category(&id, &label).await?;
            println!("Added category {} to {}", label, id);
        }
        Commands::Unlabel { id, label } => {
            let client = get_client().await?;
            client.remove_category(&id, &label).await?;
            println!("Removed category {} from {}", label, id);
        }
        Commands::Delete { id } => {
            let client = get_client().await?;
            client.trash(&id).await?;
            println!("Moved to trash {}", id);
        }
        Commands::Unsubscribe { id } => {
            let client = get_client().await?;
            let msg = client.get_message(&id).await?;
            if let Some(url) = msg.get_unsubscribe_url() {
                println!("Opening unsubscribe link: {}", url);
                open::that(&url)?;
            } else {
                anyhow::bail!("No unsubscribe link found in message headers");
            }
        }
    }

    Ok(())
}
