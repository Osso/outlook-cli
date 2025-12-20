use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};
use std::time::Duration;

const BASE_URL: &str = "https://graph.microsoft.com/v1.0";
const MAX_RETRIES: u32 = 3;
const INITIAL_BACKOFF_MS: u64 = 1000;

pub struct Client {
    http: reqwest::Client,
    access_token: String,
}

// Message list response
#[derive(Debug, Deserialize)]
pub struct MessageList {
    pub value: Option<Vec<Message>>,
    #[serde(rename = "@odata.nextLink")]
    pub next_link: Option<String>,
}

// Folder list response
#[derive(Debug, Deserialize)]
pub struct FolderList {
    pub value: Option<Vec<Folder>>,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct Folder {
    pub id: String,
    #[serde(rename = "displayName")]
    pub display_name: String,
    #[serde(rename = "parentFolderId")]
    pub parent_folder_id: Option<String>,
    #[serde(rename = "totalItemCount")]
    pub total_item_count: Option<i32>,
    #[serde(rename = "unreadItemCount")]
    pub unread_item_count: Option<i32>,
}

// Category (Outlook Master Category)
#[derive(Debug, Deserialize, Serialize)]
pub struct Category {
    pub id: Option<String>,
    #[serde(rename = "displayName")]
    pub display_name: String,
    pub color: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct CategoryList {
    pub value: Option<Vec<Category>>,
}

#[derive(Debug, Deserialize)]
pub struct Message {
    pub id: String,
    pub subject: Option<String>,
    pub from: Option<Recipient>,
    #[serde(rename = "toRecipients")]
    pub to_recipients: Option<Vec<Recipient>>,
    pub body: Option<Body>,
    #[serde(rename = "bodyPreview")]
    pub body_preview: Option<String>,
    #[serde(rename = "receivedDateTime")]
    pub received_date_time: Option<String>,
    #[serde(rename = "isRead")]
    pub is_read: Option<bool>,
    pub categories: Option<Vec<String>>,
    #[serde(rename = "internetMessageHeaders")]
    pub internet_message_headers: Option<Vec<InternetMessageHeader>>,
    #[serde(rename = "parentFolderId")]
    pub parent_folder_id: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct Recipient {
    #[serde(rename = "emailAddress")]
    pub email_address: EmailAddress,
}

#[derive(Debug, Deserialize)]
pub struct EmailAddress {
    pub name: Option<String>,
    pub address: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct Body {
    #[serde(rename = "contentType")]
    pub content_type: Option<String>,
    pub content: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct InternetMessageHeader {
    pub name: String,
    pub value: String,
}

// Move response
#[derive(Debug, Deserialize)]
pub struct MoveResponse {
    pub id: String,
}

impl Client {
    pub fn new(access_token: &str) -> Self {
        Self {
            http: reqwest::Client::builder()
                .timeout(Duration::from_secs(30))
                .build()
                .expect("Failed to build HTTP client"),
            access_token: access_token.to_string(),
        }
    }

    fn is_retryable_status(status: reqwest::StatusCode) -> bool {
        status == reqwest::StatusCode::TOO_MANY_REQUESTS
            || status.is_server_error()
            || status == reqwest::StatusCode::REQUEST_TIMEOUT
    }

    fn is_retryable_error(err: &reqwest::Error) -> bool {
        err.is_timeout() || err.is_connect() || err.is_request()
    }

    fn get_retry_delay(resp: &reqwest::Response, attempt: u32) -> Duration {
        // Check Retry-After header first (Microsoft Graph uses this for rate limits)
        if let Some(retry_after) = resp.headers().get("Retry-After") {
            if let Ok(seconds) = retry_after.to_str().unwrap_or("").parse::<u64>() {
                return Duration::from_secs(seconds);
            }
        }
        // Exponential backoff: 1s, 2s, 4s...
        Duration::from_millis(INITIAL_BACKOFF_MS * 2u64.pow(attempt))
    }

    async fn execute_with_retry<F, Fut>(&self, request_fn: F) -> Result<reqwest::Response>
    where
        F: Fn() -> Fut,
        Fut: std::future::Future<Output = Result<reqwest::Response, reqwest::Error>>,
    {
        let mut last_error = None;

        for attempt in 0..=MAX_RETRIES {
            match request_fn().await {
                Ok(resp) => {
                    if resp.status().is_success() {
                        return Ok(resp);
                    }

                    if Self::is_retryable_status(resp.status()) && attempt < MAX_RETRIES {
                        let delay = Self::get_retry_delay(&resp, attempt);
                        eprintln!(
                            "Rate limited ({}), retrying in {:?}...",
                            resp.status(),
                            delay
                        );
                        tokio::time::sleep(delay).await;
                        continue;
                    }

                    // Non-retryable error or max retries reached
                    let status = resp.status();
                    let body = resp.text().await.unwrap_or_default();
                    anyhow::bail!("HTTP {} - {}", status, body);
                }
                Err(e) => {
                    if Self::is_retryable_error(&e) && attempt < MAX_RETRIES {
                        let delay = Duration::from_millis(INITIAL_BACKOFF_MS * 2u64.pow(attempt));
                        eprintln!("Request failed ({}), retrying in {:?}...", e, delay);
                        tokio::time::sleep(delay).await;
                        last_error = Some(e);
                        continue;
                    }
                    return Err(e).context("Failed to send request");
                }
            }
        }

        Err(last_error.unwrap()).context("Failed after max retries")
    }

    async fn get<T: serde::de::DeserializeOwned>(&self, endpoint: &str) -> Result<T> {
        let url = format!("{}{}", BASE_URL, endpoint);

        let resp = self
            .execute_with_retry(|| {
                self.http
                    .get(&url)
                    .bearer_auth(&self.access_token)
                    .send()
            })
            .await?;

        resp.json().await.context("Failed to parse JSON response")
    }

    async fn post(&self, endpoint: &str) -> Result<()> {
        let url = format!("{}{}", BASE_URL, endpoint);

        self.execute_with_retry(|| {
            self.http
                .post(&url)
                .bearer_auth(&self.access_token)
                .send()
        })
        .await?;

        Ok(())
    }

    async fn post_json<T: Serialize + Sync>(&self, endpoint: &str, body: &T) -> Result<()> {
        let url = format!("{}{}", BASE_URL, endpoint);

        self.execute_with_retry(|| {
            self.http
                .post(&url)
                .bearer_auth(&self.access_token)
                .json(body)
                .send()
        })
        .await?;

        Ok(())
    }

    async fn post_json_with_response<T: Serialize + Sync, R: serde::de::DeserializeOwned>(
        &self,
        endpoint: &str,
        body: &T,
    ) -> Result<R> {
        let url = format!("{}{}", BASE_URL, endpoint);

        let resp = self
            .execute_with_retry(|| {
                self.http
                    .post(&url)
                    .bearer_auth(&self.access_token)
                    .json(body)
                    .send()
            })
            .await?;

        resp.json().await.context("Failed to parse JSON response")
    }

    async fn patch_json<T: Serialize + Sync>(&self, endpoint: &str, body: &T) -> Result<()> {
        let url = format!("{}{}", BASE_URL, endpoint);

        self.execute_with_retry(|| {
            self.http
                .patch(&url)
                .bearer_auth(&self.access_token)
                .json(body)
                .send()
        })
        .await?;

        Ok(())
    }

    // List mail folders
    pub async fn list_folders(&self) -> Result<FolderList> {
        self.get("/me/mailFolders?$top=100").await
    }

    // Get folder by well-known name or ID
    pub async fn get_folder(&self, name_or_id: &str) -> Result<Folder> {
        self.get(&format!("/me/mailFolders/{}", urlencoding::encode(name_or_id))).await
    }

    // List categories (Outlook master categories)
    pub async fn list_categories(&self) -> Result<CategoryList> {
        self.get("/me/outlook/masterCategories").await
    }

    // List messages in a folder
    pub async fn list_messages(&self, folder: &str, filter: Option<&str>, max_results: u32) -> Result<MessageList> {
        let mut endpoint = format!(
            "/me/mailFolders/{}/messages?$top={}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,categories",
            urlencoding::encode(folder),
            max_results
        );

        if let Some(f) = filter {
            endpoint.push_str(&format!("&$filter={}", urlencoding::encode(f)));
        }

        self.get(&endpoint).await
    }

    // Search messages across all folders
    pub async fn search_messages(&self, query: &str, max_results: u32) -> Result<MessageList> {
        let endpoint = format!(
            "/me/messages?$search=\"{}\"&$top={}&$select=id,subject,from,receivedDateTime,bodyPreview,isRead,categories",
            urlencoding::encode(query),
            max_results
        );
        self.get(&endpoint).await
    }

    // Get a specific message with full body and headers
    pub async fn get_message(&self, id: &str) -> Result<Message> {
        self.get(&format!(
            "/me/messages/{}?$select=id,subject,from,toRecipients,body,bodyPreview,receivedDateTime,isRead,categories,internetMessageHeaders,parentFolderId",
            urlencoding::encode(id)
        )).await
    }

    // Move message to a folder
    pub async fn move_message(&self, id: &str, destination_folder: &str) -> Result<MoveResponse> {
        let body = serde_json::json!({
            "destinationId": destination_folder
        });
        self.post_json_with_response(
            &format!("/me/messages/{}/move", urlencoding::encode(id)),
            &body,
        ).await
    }

    // Archive message (move to archive folder)
    pub async fn archive(&self, id: &str) -> Result<()> {
        self.move_message(id, "archive").await?;
        Ok(())
    }

    // Mark as spam (move to junk folder)
    pub async fn mark_spam(&self, id: &str) -> Result<()> {
        self.move_message(id, "junkemail").await?;
        Ok(())
    }

    // Unspam (move from junk to inbox)
    pub async fn unspam(&self, id: &str) -> Result<()> {
        self.move_message(id, "inbox").await?;
        Ok(())
    }

    // Move to trash (deleted items)
    pub async fn trash(&self, id: &str) -> Result<()> {
        self.move_message(id, "deleteditems").await?;
        Ok(())
    }

    // Update message categories
    pub async fn update_categories(&self, id: &str, categories: &[String]) -> Result<()> {
        let body = serde_json::json!({
            "categories": categories
        });
        self.patch_json(&format!("/me/messages/{}", urlencoding::encode(id)), &body).await
    }

    // Add a category to a message
    pub async fn add_category(&self, id: &str, category: &str) -> Result<()> {
        let msg = self.get_message(id).await?;
        let mut categories = msg.categories.unwrap_or_default();
        if !categories.iter().any(|c| c.eq_ignore_ascii_case(category)) {
            categories.push(category.to_string());
            self.update_categories(id, &categories).await?;
        }
        Ok(())
    }

    // Remove a category from a message
    pub async fn remove_category(&self, id: &str, category: &str) -> Result<()> {
        let msg = self.get_message(id).await?;
        let categories: Vec<String> = msg
            .categories
            .unwrap_or_default()
            .into_iter()
            .filter(|c| !c.eq_ignore_ascii_case(category))
            .collect();
        self.update_categories(id, &categories).await
    }

    // Mark message as read
    pub async fn mark_read(&self, id: &str) -> Result<()> {
        let body = serde_json::json!({ "isRead": true });
        self.patch_json(&format!("/me/messages/{}", urlencoding::encode(id)), &body).await
    }

    // Mark message as unread
    pub async fn mark_unread(&self, id: &str) -> Result<()> {
        let body = serde_json::json!({ "isRead": false });
        self.patch_json(&format!("/me/messages/{}", urlencoding::encode(id)), &body).await
    }
}

impl Message {
    pub fn get_from(&self) -> Option<String> {
        self.from.as_ref().and_then(|r| {
            let addr = r.email_address.address.as_deref()?;
            match r.email_address.name.as_deref() {
                Some(name) if !name.is_empty() => Some(format!("{} <{}>", name, addr)),
                _ => Some(addr.to_string()),
            }
        })
    }

    pub fn get_to(&self) -> Option<String> {
        self.to_recipients.as_ref().map(|recipients| {
            recipients
                .iter()
                .filter_map(|r| r.email_address.address.as_deref())
                .collect::<Vec<_>>()
                .join(", ")
        })
    }

    pub fn get_body_text(&self) -> Option<String> {
        self.body.as_ref().and_then(|b| b.content.clone())
    }

    pub fn get_header(&self, name: &str) -> Option<&str> {
        self.internet_message_headers.as_ref()?.iter()
            .find(|h| h.name.eq_ignore_ascii_case(name))
            .map(|h| h.value.as_str())
    }

    pub fn get_unsubscribe_url(&self) -> Option<String> {
        let header = self.get_header("List-Unsubscribe")?;
        // Parse List-Unsubscribe header, format: <http://...>, <mailto:...>
        // Prefer http/https URLs over mailto
        for part in header.split(',') {
            let part = part.trim();
            if part.starts_with('<') && part.ends_with('>') {
                let url = &part[1..part.len()-1];
                if url.starts_with("http://") || url.starts_with("https://") {
                    return Some(url.to_string());
                }
            }
        }
        // Fallback to first URL if no http found
        for part in header.split(',') {
            let part = part.trim();
            if part.starts_with('<') && part.ends_with('>') {
                return Some(part[1..part.len()-1].to_string());
            }
        }
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn make_message(from: Option<Recipient>, body: Option<Body>) -> Message {
        Message {
            id: "test123".to_string(),
            subject: Some("Test Subject".to_string()),
            from,
            to_recipients: None,
            body,
            body_preview: Some("preview".to_string()),
            received_date_time: None,
            is_read: Some(false),
            categories: None,
            internet_message_headers: None,
            parent_folder_id: None,
        }
    }

    #[test]
    fn test_get_from_with_name() {
        let msg = make_message(
            Some(Recipient {
                email_address: EmailAddress {
                    name: Some("John Doe".to_string()),
                    address: Some("john@example.com".to_string()),
                },
            }),
            None,
        );
        assert_eq!(msg.get_from(), Some("John Doe <john@example.com>".to_string()));
    }

    #[test]
    fn test_get_from_without_name() {
        let msg = make_message(
            Some(Recipient {
                email_address: EmailAddress {
                    name: None,
                    address: Some("john@example.com".to_string()),
                },
            }),
            None,
        );
        assert_eq!(msg.get_from(), Some("john@example.com".to_string()));
    }

    #[test]
    fn test_get_body_text() {
        let msg = make_message(
            None,
            Some(Body {
                content_type: Some("text/plain".to_string()),
                content: Some("Hello world".to_string()),
            }),
        );
        assert_eq!(msg.get_body_text(), Some("Hello world".to_string()));
    }

    #[test]
    fn test_get_unsubscribe_url() {
        let mut msg = make_message(None, None);
        msg.internet_message_headers = Some(vec![
            InternetMessageHeader {
                name: "List-Unsubscribe".to_string(),
                value: "<mailto:unsub@example.com>, <https://example.com/unsub>".to_string(),
            },
        ]);
        assert_eq!(msg.get_unsubscribe_url(), Some("https://example.com/unsub".to_string()));
    }

    #[test]
    fn test_get_unsubscribe_url_mailto_only() {
        let mut msg = make_message(None, None);
        msg.internet_message_headers = Some(vec![
            InternetMessageHeader {
                name: "List-Unsubscribe".to_string(),
                value: "<mailto:unsub@example.com>".to_string(),
            },
        ]);
        assert_eq!(msg.get_unsubscribe_url(), Some("mailto:unsub@example.com".to_string()));
    }
}
