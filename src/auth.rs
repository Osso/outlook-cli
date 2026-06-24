use anyhow::{Context, Result};
use oauth2::basic::BasicClient;
use oauth2::{
    AuthUrl, AuthorizationCode, ClientId, CsrfToken, PkceCodeChallenge, RedirectUrl, RefreshToken,
    Scope, TokenResponse, TokenUrl,
};
use serde::Deserialize;
use std::io::{BufRead, BufReader, Write};
use std::net::{TcpListener, TcpStream};
use std::time::{Duration, Instant};
use url::Url;

use crate::config::{self, Tokens};

// Microsoft identity platform endpoints (common = any Azure AD or personal Microsoft account)
const AUTH_URL: &str = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize";
const TOKEN_URL: &str = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
const DEVICE_CODE_URL: &str = "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode";
const LOGIN_MAX_RETRIES: u32 = 3;
const CALLBACK_TIMEOUT_SECS: u64 = 120;
const DEVICE_CODE_SCOPES: &str =
    "Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite offline_access";
const CALLBACK_POLL_INTERVAL_MS: u64 = 100;
const SLOW_DOWN_DELAY_SECS: u64 = 5;

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

#[cfg_attr(coverage_nightly, coverage(off))]
fn create_http_client() -> reqwest::Client {
    reqwest::Client::builder()
        .redirect(reqwest::redirect::Policy::none())
        .build()
        .expect("Client should build")
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn bind_callback_listener() -> Result<TcpListener> {
    // Bind to port 0 to get an OS-assigned available port (prevents port squatting)
    let listener = TcpListener::bind("127.0.0.1:0").context("Failed to bind to local port")?;
    listener.set_nonblocking(true)?;
    Ok(listener)
}

fn build_tokens(access_token: String, refresh_token: Option<String>) -> Result<Tokens> {
    let refresh_token =
        refresh_token.ok_or_else(|| anyhow::anyhow!("No refresh token received"))?;
    Ok(Tokens {
        access_token,
        refresh_token,
    })
}

#[cfg_attr(coverage_nightly, coverage(off))]
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

#[cfg_attr(coverage_nightly, coverage(off))]
async fn try_login(client_id: &str) -> Result<Tokens> {
    let listener = bind_callback_listener()?;
    let port = listener.local_addr()?.port();
    let client = build_oauth_client(client_id, port)?;
    let http_client = create_http_client();
    let (pkce_challenge, pkce_verifier) = PkceCodeChallenge::new_random_sha256();
    let (auth_url, csrf_token) = build_authorization_url(&client, pkce_challenge);

    println!("Opening browser for authentication...");
    open::that(auth_url.as_str())?;
    let code = wait_for_callback_with_timeout(listener, csrf_token, CALLBACK_TIMEOUT_SECS)?;
    let token_result = client
        .exchange_code(code)
        .set_pkce_verifier(pkce_verifier)
        .request_async(&http_client)
        .await
        .context("Failed to exchange code for token")?;

    let tokens = build_tokens(
        token_result.access_token().secret().to_string(),
        token_result
            .refresh_token()
            .map(|token| token.secret().to_string()),
    )?;
    config::save_tokens(&tokens)?;
    Ok(tokens)
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn build_oauth_client(
    client_id: &str,
    port: u16,
) -> Result<
    BasicClient<
        oauth2::EndpointSet,
        oauth2::EndpointNotSet,
        oauth2::EndpointNotSet,
        oauth2::EndpointNotSet,
        oauth2::EndpointSet,
    >,
> {
    Ok(BasicClient::new(ClientId::new(client_id.to_string()))
        .set_auth_uri(AuthUrl::new(AUTH_URL.to_string())?)
        .set_token_uri(TokenUrl::new(TOKEN_URL.to_string())?)
        .set_redirect_uri(RedirectUrl::new(format!("http://localhost:{}", port))?))
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn build_authorization_url(
    client: &BasicClient<
        oauth2::EndpointSet,
        oauth2::EndpointNotSet,
        oauth2::EndpointNotSet,
        oauth2::EndpointNotSet,
        oauth2::EndpointSet,
    >,
    pkce_challenge: PkceCodeChallenge,
) -> (url::Url, CsrfToken) {
    client
        .authorize_url(CsrfToken::new_random)
        .add_scope(Scope::new("Mail.ReadWrite".to_string()))
        .add_scope(Scope::new("Mail.Send".to_string()))
        .add_scope(Scope::new("MailboxSettings.ReadWrite".to_string()))
        .add_scope(Scope::new("offline_access".to_string()))
        .set_pkce_challenge(pkce_challenge)
        .url()
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn wait_for_connection(listener: &TcpListener, deadline: Instant) -> Result<TcpStream> {
    loop {
        match listener.accept() {
            Ok((stream, _)) => return Ok(stream),
            Err(ref err) if err.kind() == std::io::ErrorKind::WouldBlock => {
                if Instant::now() >= deadline {
                    anyhow::bail!("Timeout waiting for OAuth callback");
                }
                std::thread::sleep(Duration::from_millis(CALLBACK_POLL_INTERVAL_MS));
            }
            Err(err) => return Err(err).context("Failed to accept connection"),
        }
    }
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn read_request_line(stream: &TcpStream) -> Result<String> {
    let mut reader = BufReader::new(stream);
    let mut request_line = String::new();
    reader.read_line(&mut request_line)?;
    Ok(request_line)
}

fn extract_redirect_url(request_line: &str) -> Result<&str> {
    request_line
        .split_whitespace()
        .nth(1)
        .ok_or_else(|| anyhow::anyhow!("Invalid request"))
}

fn query_value(url: &Url, key: &str) -> Option<String> {
    url.query_pairs()
        .find(|(query_key, _)| query_key == key)
        .map(|(_, value)| value.into_owned())
}

fn extract_callback_code(url: &Url) -> Result<AuthorizationCode> {
    let code = query_value(url, "code").ok_or_else(|| anyhow::anyhow!("No code in callback"))?;
    Ok(AuthorizationCode::new(code))
}

fn verify_callback_state(url: &Url, expected_csrf: &CsrfToken) -> Result<()> {
    let state = query_value(url, "state").ok_or_else(|| anyhow::anyhow!("No state in callback"))?;
    if state == expected_csrf.secret().as_str() {
        return Ok(());
    }
    anyhow::bail!("CSRF token mismatch")
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn write_success_response(stream: &mut TcpStream) -> Result<()> {
    let response = "HTTP/1.1 200 OK\r\nContent-Type: text/html\r\n\r\n<html><body><h1>Authentication successful!</h1><p>You can close this window.</p></body></html>";
    stream.write_all(response.as_bytes())?;
    Ok(())
}

#[cfg_attr(coverage_nightly, coverage(off))]
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

    let deadline = Instant::now() + Duration::from_secs(timeout_secs);
    let mut stream = wait_for_connection(&listener, deadline)?;
    stream.set_nonblocking(false)?;

    let request_line = read_request_line(&stream)?;
    let redirect_url = extract_redirect_url(&request_line)?;
    let url = Url::parse(&format!("http://localhost{}", redirect_url))?;
    let code = extract_callback_code(&url)?;
    verify_callback_state(&url, &expected_csrf)?;
    write_success_response(&mut stream)?;
    Ok(code)
}

#[cfg_attr(coverage_nightly, coverage(off))]
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
#[cfg_attr(coverage_nightly, coverage(off))]
pub async fn login_device_code(client_id: &str) -> Result<Tokens> {
    let http_client = create_http_client();
    let device_response = request_device_code(&http_client, client_id, DEVICE_CODE_SCOPES).await?;
    print_device_login_instructions(&device_response);

    let _ = open::that(&device_response.verification_uri);

    let tokens = wait_for_device_tokens(&http_client, client_id, &device_response).await?;
    config::save_tokens(&tokens)?;
    println!("Authentication successful!");
    Ok(tokens)
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn request_device_code(
    http_client: &reqwest::Client,
    client_id: &str,
    scopes: &str,
) -> Result<DeviceCodeResponse> {
    http_client
        .post(DEVICE_CODE_URL)
        .form(&[("client_id", client_id), ("scope", scopes)])
        .send()
        .await
        .context("Failed to request device code")?
        .json::<DeviceCodeResponse>()
        .await
        .context("Failed to parse device code response")
}

#[cfg_attr(coverage_nightly, coverage(off))]
fn print_device_login_instructions(device_response: &DeviceCodeResponse) {
    println!("\nTo sign in, open: {}", device_response.verification_uri);
    println!("Enter code: {}\n", device_response.user_code);
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn poll_device_token_body(
    http_client: &reqwest::Client,
    client_id: &str,
    device_code: &str,
) -> Result<String> {
    let response = http_client
        .post(TOKEN_URL)
        .form(&[
            ("client_id", client_id),
            ("device_code", device_code),
            ("grant_type", "urn:ietf:params:oauth:grant-type:device_code"),
        ])
        .send()
        .await
        .context("Failed to poll for token")?;

    response.text().await.map_err(Into::into)
}

enum DevicePollResult {
    Authorized(Tokens),
    Pending,
    SlowDown,
}

fn parse_device_poll_result(body: &str) -> Result<DevicePollResult> {
    if let Ok(token_response) = serde_json::from_str::<DeviceTokenResponse>(body) {
        return Ok(DevicePollResult::Authorized(build_tokens(
            token_response.access_token,
            token_response.refresh_token,
        )?));
    }

    if let Ok(error) = serde_json::from_str::<DeviceTokenError>(body) {
        return match error.error.as_str() {
            "authorization_pending" => Ok(DevicePollResult::Pending),
            "slow_down" => Ok(DevicePollResult::SlowDown),
            _ => anyhow::bail!("Authentication failed: {}", error.error),
        };
    }

    anyhow::bail!("Authentication failed: unrecognized device token response")
}

#[cfg_attr(coverage_nightly, coverage(off))]
async fn wait_for_device_tokens(
    http_client: &reqwest::Client,
    client_id: &str,
    device_response: &DeviceCodeResponse,
) -> Result<Tokens> {
    let deadline = Instant::now() + Duration::from_secs(device_response.expires_in);
    let interval = Duration::from_secs(device_response.interval);

    loop {
        if Instant::now() >= deadline {
            anyhow::bail!("Device code expired");
        }

        tokio::time::sleep(interval).await;
        let body =
            poll_device_token_body(http_client, client_id, &device_response.device_code).await?;

        match parse_device_poll_result(&body)? {
            DevicePollResult::Authorized(tokens) => return Ok(tokens),
            DevicePollResult::Pending => continue,
            DevicePollResult::SlowDown => {
                tokio::time::sleep(Duration::from_secs(SLOW_DOWN_DELAY_SECS)).await;
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn build_tokens_requires_refresh_token() {
        let err = build_tokens("access".to_string(), None).unwrap_err();

        assert!(err.to_string().contains("No refresh token"));
    }

    #[test]
    fn extract_redirect_url_parses_get_request_target() {
        let redirect = extract_redirect_url("GET /callback?code=abc&state=xyz HTTP/1.1").unwrap();

        assert_eq!(redirect, "/callback?code=abc&state=xyz");
    }

    #[test]
    fn extract_redirect_url_rejects_malformed_request_line() {
        assert!(extract_redirect_url("GET").is_err());
    }

    #[test]
    fn query_value_extracts_owned_value() {
        let url = Url::parse("http://localhost/callback?code=abc&state=xyz").unwrap();

        assert_eq!(query_value(&url, "code").as_deref(), Some("abc"));
        assert_eq!(query_value(&url, "missing"), None);
    }

    #[test]
    fn extract_callback_code_requires_code_query_param() {
        let url = Url::parse("http://localhost/callback?code=abc").unwrap();
        let code = extract_callback_code(&url).unwrap();

        assert_eq!(code.secret(), "abc");
        assert!(extract_callback_code(&Url::parse("http://localhost/callback").unwrap()).is_err());
    }

    #[test]
    fn verify_callback_state_checks_expected_csrf() {
        let expected = CsrfToken::new("state-1".to_string());
        let good = Url::parse("http://localhost/callback?state=state-1").unwrap();
        let bad = Url::parse("http://localhost/callback?state=state-2").unwrap();

        assert!(verify_callback_state(&good, &expected).is_ok());
        assert!(verify_callback_state(&bad, &expected).is_err());
    }

    #[test]
    fn parse_device_poll_result_handles_success_pending_and_slow_down() {
        let success = r#"{"access_token":"access","refresh_token":"refresh"}"#;
        let pending = r#"{"error":"authorization_pending"}"#;
        let slow_down = r#"{"error":"slow_down"}"#;

        match parse_device_poll_result(success).unwrap() {
            DevicePollResult::Authorized(tokens) => {
                assert_eq!(tokens.access_token, "access");
                assert_eq!(tokens.refresh_token, "refresh");
            }
            _ => panic!("expected authorized tokens"),
        }
        assert!(matches!(
            parse_device_poll_result(pending).unwrap(),
            DevicePollResult::Pending
        ));
        assert!(matches!(
            parse_device_poll_result(slow_down).unwrap(),
            DevicePollResult::SlowDown
        ));
    }

    #[test]
    fn parse_device_poll_result_reports_unknown_errors_and_invalid_json() {
        let denied = r#"{"error":"authorization_declined"}"#;

        assert!(parse_device_poll_result(denied).is_err());
        assert!(parse_device_poll_result("not json").is_err());
    }
}
