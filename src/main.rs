#![cfg_attr(coverage_nightly, feature(coverage_attribute))]

mod api;
mod auth;
mod config;

use anyhow::Result;
use clap::{Parser, Subcommand};
use std::collections::HashSet;

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
    /// Set custom OAuth client ID (optional - has built-in default)
    Config {
        /// Client ID (Application ID from Azure)
        client_id: String,
    },
    /// Authenticate with Microsoft (opens browser)
    Login {
        /// Use device code flow (for first-party app IDs that don't allow localhost redirect)
        #[arg(long, short)]
        device: bool,
    },
    /// List categories (like Gmail labels)
    Labels,
    /// Sync categories: create master categories for any used on messages
    SyncLabels,
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
    /// Clear all categories from a message
    ClearLabels {
        /// Message ID (or "all" to clear from all inbox messages)
        id: String,
    },
    /// Mark a message as read
    MarkRead {
        /// Message ID
        id: String,
    },
    /// Mark a message as unread
    MarkUnread {
        /// Message ID
        id: String,
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

#[cfg_attr(coverage_nightly, coverage(off))]
async fn get_client() -> Result<api::Client> {
    let cfg = config::load_config()?;
    let client_id = cfg.client_id();

    let tokens = match config::load_tokens() {
        Ok(t) => t,
        Err(_) => anyhow::bail!("Not logged in. Run 'outlook login' first"),
    };

    let client = api::Client::new(&tokens.access_token);

    match client.list_folders().await {
        Ok(_) => Ok(client),
        Err(_) => {
            let new_tokens = auth::refresh_token(client_id, &tokens.refresh_token).await?;
            Ok(api::Client::new(&new_tokens.access_token))
        }
    }
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn save_config(client_id: String) -> Result<()> {
    let cfg = config::Config {
        client_id: Some(client_id),
    };
    config::save_config(&cfg)?;
    println!("Custom client ID saved to {:?}", config::config_dir());
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn login(device: bool) -> Result<()> {
    let cfg = config::load_config()?;
    let client_id = cfg.client_id();

    let _ = std::fs::remove_file(config::tokens_path());

    if device {
        auth::login_device_code(client_id).await?;
    } else {
        auth::login(client_id).await?;
    }
    println!("Login successful! Tokens saved.");
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn list_labels(json: bool) -> Result<()> {
    let client = get_client().await?;
    let categories = client.list_categories().await?;

    if let Some(cats) = categories.value {
        if json {
            println!("{}", serde_json::to_string(&cats)?);
        } else {
            println!("Categories:");
            for cat in cats {
                let color = cat.color.as_deref().unwrap_or("none");
                println!("  {} (color: {})", cat.display_name, color);
            }
        }
    } else if json {
        println!("[]");
    } else {
        println!("No categories found.");
    }
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn sync_labels() -> Result<()> {
    let client = get_client().await?;
    let master_names: HashSet<String> = client
        .list_categories()
        .await?
        .value
        .unwrap_or_default()
        .into_iter()
        .map(|c| c.display_name.to_lowercase())
        .collect();

    let found = find_missing_categories(
        client.list_messages("inbox", None, 200).await?.value,
        &master_names,
    );

    if found.is_empty() {
        println!("All categories are already in master list.");
        return Ok(());
    }

    for cat in &found {
        client.create_category(cat, None).await?;
        println!("Created category: {}", cat);
    }
    println!("Synced {} categories.", found.len());
    Ok(())
}

fn find_missing_categories(
    messages: Option<Vec<api::Message>>,
    master_names: &HashSet<String>,
) -> HashSet<String> {
    messages
        .unwrap_or_default()
        .into_iter()
        .flat_map(|msg| msg.categories.unwrap_or_default().into_iter())
        .filter(|category| !master_names.contains(&category.to_lowercase()))
        .collect()
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn list_messages(
    max: u32,
    query: Option<String>,
    label: String,
    unread: bool,
    json: bool,
) -> Result<()> {
    let client = get_client().await?;
    let folder = normalize_folder(&label);
    let list = if let Some(q) = query.as_deref() {
        client.search_messages(q, max).await?
    } else {
        let filter = unread.then_some("isRead eq false");
        client.list_messages(&folder, filter, max).await?
    };

    match list.value {
        Some(messages) if json => print_messages_json(&messages)?,
        Some(messages) => print_messages_text(messages),
        None if json => println!("[]"),
        None => println!("No messages found."),
    }

    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn print_messages_json(messages: &[api::Message]) -> Result<()> {
    let items: Vec<_> = messages.iter().map(list_message_json_item).collect();
    println!("{}", serde_json::to_string(&items)?);
    Ok(())
}

fn list_message_json_item(msg: &api::Message) -> serde_json::Value {
    serde_json::json!({
        "id": msg.id,
        "from": msg.get_from(),
        "subject": msg.subject,
        "date": msg.received_date_time,
        "snippet": msg.body_preview,
        "isRead": msg.is_read,
        "categories": msg.categories,
    })
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn print_messages_text(messages: Vec<api::Message>) {
    for msg in messages {
        let from = msg.get_from().unwrap_or_else(|| "Unknown".to_string());
        let subject = msg.subject.as_deref().unwrap_or("(no subject)");
        println!("{} | {} | {}", msg.id, from, subject);
    }
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn read_message(id: String, json: bool) -> Result<()> {
    let client = get_client().await?;
    let msg = client.get_message(&id).await?;

    if json {
        return print_message_json(&msg);
    }

    println!(
        "From: {}",
        msg.get_from().unwrap_or_else(|| "Unknown".to_string())
    );
    println!(
        "To: {}",
        msg.get_to().unwrap_or_else(|| "Unknown".to_string())
    );
    println!(
        "Subject: {}",
        msg.subject.as_deref().unwrap_or("(no subject)")
    );
    println!(
        "Date: {}",
        msg.received_date_time.as_deref().unwrap_or("Unknown")
    );
    println!("---");

    if let Some(body) = msg.get_body_text() {
        println!("{}", body);
    } else if let Some(preview) = &msg.body_preview {
        println!("{}", preview);
    }
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn print_message_json(msg: &api::Message) -> Result<()> {
    println!(
        "{}",
        serde_json::to_string(&serde_json::json!({
            "id": msg.id,
            "from": msg.get_from(),
            "to": msg.get_to(),
            "subject": msg.subject,
            "date": msg.received_date_time,
            "body": msg.get_body_text(),
            "snippet": msg.body_preview,
            "isRead": msg.is_read,
            "categories": msg.categories,
        }))?
    );
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn archive_message(id: String) -> Result<()> {
    let client = get_client().await?;
    client.archive(&id).await?;
    println!("Archived {}", id);
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn spam_message(id: String) -> Result<()> {
    let client = get_client().await?;
    let msg = client.get_message(&id).await?;
    if let Some(url) = msg.get_unsubscribe_url() {
        if url.starts_with("http") {
            let _ = open::that(&url);
        }
    }
    client.mark_spam(&id).await?;
    println!("Marked as spam {}", id);
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn unspam_message(id: String) -> Result<()> {
    let client = get_client().await?;
    client.unspam(&id).await?;
    println!("Moved to inbox {}", id);
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn add_label(id: String, label: String) -> Result<()> {
    let client = get_client().await?;
    client.ensure_category(&label).await?;
    client.add_category(&id, &label).await?;
    println!("Added category {} to {}", label, id);
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn remove_label(id: String, label: String) -> Result<()> {
    let client = get_client().await?;
    client.remove_category(&id, &label).await?;
    println!("Removed category {} from {}", label, id);
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn clear_labels(id: String) -> Result<()> {
    let client = get_client().await?;
    if id == "all" {
        return clear_all_inbox_labels(&client).await;
    }

    client.update_categories(&id, &[]).await?;
    println!("Cleared all categories from {}", id);
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn clear_all_inbox_labels(client: &api::Client) -> Result<()> {
    let mut count = 0;
    let messages = client.list_messages("inbox", None, 200).await?;
    for msg in messages.value.unwrap_or_default() {
        if !message_has_categories(&msg) {
            continue;
        }

        client.update_categories(&msg.id, &[]).await?;
        println!(
            "Cleared categories from: {}",
            msg.subject.as_deref().unwrap_or("(no subject)")
        );
        count += 1;
    }

    println!("Cleared categories from {} messages.", count);
    Ok(())
}

fn message_has_categories(message: &api::Message) -> bool {
    message
        .categories
        .as_ref()
        .map(|categories| !categories.is_empty())
        .unwrap_or(false)
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn mark_read(id: String) -> Result<()> {
    let client = get_client().await?;
    client.mark_read(&id).await?;
    println!("Marked as read: {}", id);
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn mark_unread(id: String) -> Result<()> {
    let client = get_client().await?;
    client.mark_unread(&id).await?;
    println!("Marked as unread: {}", id);
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn delete_message(id: String) -> Result<()> {
    let client = get_client().await?;
    client.trash(&id).await?;
    println!("Moved to trash {}", id);
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn unsubscribe(id: String) -> Result<()> {
    let client = get_client().await?;
    let msg = client.get_message(&id).await?;
    if let Some(url) = msg.get_unsubscribe_url() {
        println!("Opening unsubscribe link: {}", url);
        open::that(&url)?;
    } else {
        anyhow::bail!("No unsubscribe link found in message headers");
    }
    Ok(())
}

#[tokio::main]
#[cfg_attr(coverage_nightly, coverage(off))]
async fn main() -> Result<()> {
    let cli = Cli::parse();
    run_command(cli.command, cli.json).await
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn run_command(command: Commands, json: bool) -> Result<()> {
    match command {
        Commands::Config { client_id } => save_config(client_id)?,
        Commands::Login { device } => login(device).await?,
        Commands::Labels => list_labels(json).await?,
        Commands::SyncLabels => sync_labels().await?,
        Commands::List {
            max,
            query,
            label,
            unread,
        } => list_messages(max, query, label, unread, json).await?,
        Commands::Read { id } => read_message(id, json).await?,
        other => run_message_action(other).await?,
    }
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn run_message_action(command: Commands) -> Result<()> {
    match command {
        Commands::Label { id, label } => add_label(id, label).await?,
        Commands::Unlabel { id, label } => remove_label(id, label).await?,
        Commands::ClearLabels { id } => clear_labels(id).await?,
        other => run_single_id_action(other).await?,
    }
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn run_single_id_action(command: Commands) -> Result<()> {
    match command {
        Commands::Archive { id } => archive_message(id).await?,
        Commands::Spam { id } => spam_message(id).await?,
        Commands::Unspam { id } => unspam_message(id).await?,
        Commands::MarkRead { id } => mark_read(id).await?,
        Commands::MarkUnread { id } => mark_unread(id).await?,
        Commands::Delete { id } => delete_message(id).await?,
        Commands::Unsubscribe { id } => unsubscribe(id).await?,
        _ => unreachable!("Command should be handled by run_command"),
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use clap::CommandFactory;

    fn message_with_categories(categories: Option<Vec<&str>>) -> api::Message {
        api::Message {
            id: "id-1".to_string(),
            subject: Some("Subject".to_string()),
            from: None,
            to_recipients: None,
            body: None,
            body_preview: Some("Preview".to_string()),
            received_date_time: Some("2026-06-24T12:00:00Z".to_string()),
            is_read: Some(false),
            categories: categories.map(|items| items.into_iter().map(str::to_string).collect()),
            internet_message_headers: None,
            parent_folder_id: None,
        }
    }

    #[test]
    fn cli_definition_is_valid() {
        Cli::command().debug_assert();
    }

    #[test]
    fn normalize_folder_maps_common_aliases() {
        assert_eq!(normalize_folder("Sent"), "sentitems");
        assert_eq!(normalize_folder("deleted"), "deleteditems");
        assert_eq!(normalize_folder("junk"), "junkemail");
        assert_eq!(normalize_folder("custom"), "custom");
    }

    #[test]
    fn find_missing_categories_returns_categories_absent_from_master_list() {
        let messages = Some(vec![
            message_with_categories(Some(vec!["Travel", "Receipts"])),
            message_with_categories(Some(vec!["travel", "Followup"])),
        ]);
        let master_names = HashSet::from(["travel".to_string()]);

        let missing = find_missing_categories(messages, &master_names);

        assert!(missing.contains("Receipts"));
        assert!(missing.contains("Followup"));
        assert!(!missing.contains("Travel"));
    }

    #[test]
    fn list_message_json_item_contains_expected_summary_fields() {
        let msg = message_with_categories(Some(vec!["Receipts"]));

        let item = list_message_json_item(&msg);

        assert_eq!(item["id"], "id-1");
        assert_eq!(item["subject"], "Subject");
        assert_eq!(item["snippet"], "Preview");
        assert_eq!(item["isRead"], false);
        assert_eq!(item["categories"][0], "Receipts");
    }

    #[test]
    fn message_has_categories_requires_non_empty_categories() {
        assert!(message_has_categories(&message_with_categories(Some(
            vec!["Receipts"]
        ))));
        assert!(!message_has_categories(&message_with_categories(Some(
            vec![]
        ))));
        assert!(!message_has_categories(&message_with_categories(None)));
    }
}
