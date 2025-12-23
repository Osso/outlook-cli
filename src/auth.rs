use anyhow::{Context, Result};
use oauth2::basic::BasicClient;
use oauth2::{
    AuthUrl, AuthorizationCode, ClientId, CsrfToken, PkceCodeChallenge, RedirectUrl, RefreshToken,
    Scope, TokenResponse, TokenUrl,
};
use serde::Deserialize;
use std::io::{BufRead, BufReader, Write};
use std::net::TcpListener;
use std::time::Duration;
use url::Url;

use crate::config::{self, Tokens};

// Microsoft identity platform endpoints (common = any Azure AD or personal Microsoft account)
const AUTH_URL: &str = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
const TOKEN_URL: &str = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
const DEVICE_CODE_URL: &str = "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode";
const LOGIN_MAX_RETRIES: u32 = 3;
const CALLBACK_TIMEOUT_SECS: u64 = 120;

#[derive(Deserialize)]
struct DeviceCodeResponse {
    device_code: String,
    user_code: String,
    verification_uri: String,
    expires_in: u64,
    interval: u64,
}

#[derive(Deserialize)]
struct DeviceTokenResponse {
    access_token: String,
    refresh_token: Option<String>,
}

#[derive(Deserialize)]
struct DeviceTokenError {
    error: String,
}

fn create_http_client() -> reqwest::Client {
    reqwest::Client::builder()
        .redirect(reqwest::redirect::Policy::none())
        .build()
        .expect("Client should build")
}

pub async fn login(client_id: &str) -> Result<Tokens> {
    let mut last_error = None;

    for attempt in 0..LOGIN_MAX_RETRIES {
        if attempt > 0 {
            eprintln!(
                "Retrying login (attempt {}/{})...",
                attempt + 1,
                LOGIN_MAX_RETRIES
            );
        }

        match try_login(client_id).await {
            Ok(tokens) => return Ok(tokens),
            Err(e) => {
                eprintln!("Login failed: {}", e);
                last_error = Some(e);
            }
        }
    }

    Err(last_error
        .unwrap_or_else(|| anyhow::anyhow!("Login failed after {} attempts", LOGIN_MAX_RETRIES)))
}

async fn try_login(client_id: &str) -> Result<Tokens> {
    // Bind to port 0 to get an OS-assigned available port (prevents port squatting)
    let listener = TcpListener::bind("127.0.0.1:0").context("Failed to bind to local port")?;
    let port = listener.local_addr()?.port();

    // Set timeout on listener so we don't wait forever
    listener.set_nonblocking(true)?;

    // Public client - no client_secret needed, PKCE provides security
    let client = BasicClient::new(ClientId::new(client_id.to_string()))
        .set_auth_uri(AuthUrl::new(AUTH_URL.to_string())?)
        .set_token_uri(TokenUrl::new(TOKEN_URL.to_string())?)
        .set_redirect_uri(RedirectUrl::new(format!("http://localhost:{}", port))?);

    let http_client = create_http_client();

    let (pkce_challenge, pkce_verifier) = PkceCodeChallenge::new_random_sha256();

    let (auth_url, csrf_token) = client
        .authorize_url(CsrfToken::new_random)
        // Mail.ReadWrite for reading/moving/deleting messages
        // Mail.Send for potential future send support
        // MailboxSettings.ReadWrite for reading and creating categories
        // offline_access for refresh token
        .add_scope(Scope::new("Mail.ReadWrite".to_string()))
        .add_scope(Scope::new("Mail.Send".to_string()))
        .add_scope(Scope::new("MailboxSettings.ReadWrite".to_string()))
        .add_scope(Scope::new("offline_access".to_string()))
        .set_pkce_challenge(pkce_challenge)
        .url();

    let url = auth_url.to_string();

    println!("Opening browser for authentication...");
    open::that(&url)?;

    let code = wait_for_callback_with_timeout(listener, csrf_token, CALLBACK_TIMEOUT_SECS)?;

    let token_result = client
        .exchange_code(code)
        .set_pkce_verifier(pkce_verifier)
        .request_async(&http_client)
        .await
        .context("Failed to exchange code for token")?;

    let tokens = Tokens {
        access_token: token_result.access_token().secret().to_string(),
        refresh_token: token_result
            .refresh_token()
            .map(|t| t.secret().to_string())
            .ok_or_else(|| anyhow::anyhow!("No refresh token received"))?,
    };

    config::save_tokens(&tokens)?;
    Ok(tokens)
}

fn wait_for_callback_with_timeout(
    listener: TcpListener,
    expected_csrf: CsrfToken,
    timeout_secs: u64,
) -> Result<AuthorizationCode> {
    let port = listener.local_addr()?.port();
    println!(
        "Waiting for OAuth callback on port {} (timeout: {}s)...",
        port, timeout_secs
    );

    let deadline = std::time::Instant::now() + Duration::from_secs(timeout_secs);

    // Poll for connection with timeout
    let (mut stream, _) = loop {
        match listener.accept() {
            Ok(conn) => break conn,
            Err(ref e) if e.kind() == std::io::ErrorKind::WouldBlock => {
                if std::time::Instant::now() >= deadline {
                    anyhow::bail!("Timeout waiting for OAuth callback");
                }
                std::thread::sleep(Duration::from_millis(100));
                continue;
            }
            Err(e) => return Err(e).context("Failed to accept connection"),
        }
    };

    // Set stream to blocking for reading
    stream.set_nonblocking(false)?;

    let mut reader = BufReader::new(&stream);
    let mut request_line = String::new();
    reader.read_line(&mut request_line)?;

    let redirect_url = request_line
        .split_whitespace()
        .nth(1)
        .ok_or_else(|| anyhow::anyhow!("Invalid request"))?;

    let url = Url::parse(&format!("http://localhost{}", redirect_url))?;

    let code = url
        .query_pairs()
        .find(|(key, _)| key == "code")
        .map(|(_, value)| AuthorizationCode::new(value.into_owned()))
        .ok_or_else(|| anyhow::anyhow!("No code in callback"))?;

    let state = url
        .query_pairs()
        .find(|(key, _)| key == "state")
        .map(|(_, value)| CsrfToken::new(value.into_owned()))
        .ok_or_else(|| anyhow::anyhow!("No state in callback"))?;

    if state.secret() != expected_csrf.secret() {
        anyhow::bail!("CSRF token mismatch");
    }

    let response = "HTTP/1.1 200 OK\r\nContent-Type: text/html\r\n\r\n<html><body><h1>Authentication successful!</h1><p>You can close this window.</p></body></html>";
    stream.write_all(response.as_bytes())?;

    Ok(code)
}

pub async fn refresh_token(client_id: &str, refresh: &str) -> Result<Tokens> {
    // Public client - no client_secret needed
    let client = BasicClient::new(ClientId::new(client_id.to_string()))
        .set_auth_uri(AuthUrl::new(AUTH_URL.to_string())?)
        .set_token_uri(TokenUrl::new(TOKEN_URL.to_string())?);

    let http_client = create_http_client();

    let token_result = client
        .exchange_refresh_token(&RefreshToken::new(refresh.to_string()))
        .request_async(&http_client)
        .await
        .context("Failed to refresh token")?;

    let tokens = Tokens {
        access_token: token_result.access_token().secret().to_string(),
        refresh_token: token_result
            .refresh_token()
            .map(|t| t.secret().to_string())
            .unwrap_or_else(|| refresh.to_string()),
    };

    config::save_tokens(&tokens)?;
    Ok(tokens)
}

/// Device code flow - works with first-party Microsoft app IDs without redirect URI
pub async fn login_device_code(client_id: &str) -> Result<Tokens> {
    let http_client = create_http_client();
    let scopes = "Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite offline_access";

    // Step 1: Request device code
    let device_response = http_client
        .post(DEVICE_CODE_URL)
        .form(&[("client_id", client_id), ("scope", scopes)])
        .send()
        .await
        .context("Failed to request device code")?
        .json::<DeviceCodeResponse>()
        .await
        .context("Failed to parse device code response")?;

    println!("\nTo sign in, open: {}", device_response.verification_uri);
    println!("Enter code: {}\n", device_response.user_code);

    // Try to open browser
    let _ = open::that(&device_response.verification_uri);

    // Step 2: Poll for token
    let deadline = std::time::Instant::now() + Duration::from_secs(device_response.expires_in);
    let interval = Duration::from_secs(device_response.interval);

    loop {
        if std::time::Instant::now() >= deadline {
            anyhow::bail!("Device code expired");
        }

        tokio::time::sleep(interval).await;

        let response = http_client
            .post(TOKEN_URL)
            .form(&[
                ("client_id", client_id),
                ("device_code", &device_response.device_code),
                ("grant_type", "urn:ietf:params:oauth:grant-type:device_code"),
            ])
            .send()
            .await
            .context("Failed to poll for token")?;

        let body = response.text().await?;

        // Try to parse as success
        if let Ok(token_response) = serde_json::from_str::<DeviceTokenResponse>(&body) {
            let tokens = Tokens {
                access_token: token_response.access_token,
                refresh_token: token_response
                    .refresh_token
                    .ok_or_else(|| anyhow::anyhow!("No refresh token received"))?,
            };
            config::save_tokens(&tokens)?;
            println!("Authentication successful!");
            return Ok(tokens);
        }

        // Check if still pending
        if let Ok(error) = serde_json::from_str::<DeviceTokenError>(&body) {
            match error.error.as_str() {
                "authorization_pending" => continue,
                "slow_down" => {
                    tokio::time::sleep(Duration::from_secs(5)).await;
                    continue;
                }
                _ => anyhow::bail!("Authentication failed: {}", error.error),
            }
        }
    }
}
