//! Where MCP tool calls end up: a control-channel method call.

use std::io::{BufRead, BufReader, Write};
use std::net::{TcpStream, ToSocketAddrs};
use std::time::Duration;

use serde_json::{Value, json};

/// Something that answers control-channel methods (`engine.execute`, `document.inspect`,
/// `ui.pointer`, `ui.render`, …). See `deckcraft_ui_egui::control` for the full list.
pub trait Backend {
    /// Call one method. `Ok` carries the `result`, `Err` the error message.
    fn call(&mut self, method: &str, params: Value) -> Result<Value, String>;
    /// True when a real UI is attached (`ui.screenshot`, `ui.click`, dialogs… work).
    fn has_ui(&self) -> bool;
    /// Short human description ("headless", "connected to 127.0.0.1:7979").
    fn describe(&self) -> String;
}

/// A running DeckCraft app, reached through its loopback control port.
pub struct Remote {
    addr: String,
    conn: Option<(BufReader<TcpStream>, TcpStream)>,
    next_id: u64,
}

impl Remote {
    /// Connect to `addr` (`127.0.0.1:7979`), failing fast when nothing is listening.
    pub fn connect(addr: &str) -> std::io::Result<Self> {
        let mut r = Self { addr: addr.to_string(), conn: None, next_id: 1 };
        r.reconnect()?;
        Ok(r)
    }

    pub fn addr(&self) -> &str {
        &self.addr
    }

    fn reconnect(&mut self) -> std::io::Result<()> {
        self.conn = None;
        let mut last = std::io::Error::new(std::io::ErrorKind::NotFound, format!("cannot resolve {}", self.addr));
        for sa in self.addr.to_socket_addrs()? {
            match TcpStream::connect_timeout(&sa, Duration::from_millis(800)) {
                Ok(s) => {
                    s.set_nodelay(true).ok();
                    // The app answers within 60 s (its own timeout); leave headroom.
                    s.set_read_timeout(Some(Duration::from_secs(90))).ok();
                    let read = s.try_clone()?;
                    self.conn = Some((BufReader::new(read), s));
                    return Ok(());
                }
                Err(e) => last = e,
            }
        }
        Err(last)
    }

    /// Write one request line. A failure here means the app never got a complete line, so retrying is safe.
    /// A connection the app has already closed (it restarted since the last call) is replaced before writing.
    fn send(&mut self, line: &str) -> std::io::Result<()> {
        if self.conn.as_ref().is_none_or(|(_, writer)| is_closed(writer)) {
            self.reconnect()?;
        }
        let Some((_, writer)) = self.conn.as_mut() else {
            return Err(std::io::Error::new(std::io::ErrorKind::NotConnected, "not connected"));
        };
        writer.write_all(format!("{line}\n").as_bytes())?;
        writer.flush()
    }

    /// Read the reply to a request that was already sent.
    fn receive(&mut self) -> std::io::Result<String> {
        let Some((reader, _)) = self.conn.as_mut() else {
            return Err(std::io::Error::new(std::io::ErrorKind::NotConnected, "not connected"));
        };
        let mut reply = String::new();
        if reader.read_line(&mut reply)? == 0 {
            return Err(std::io::Error::new(std::io::ErrorKind::UnexpectedEof, "control channel closed"));
        }
        Ok(reply)
    }
}

/// True when the peer has closed or reset `stream`, or sent bytes nobody asked for. Between calls nothing
/// should be waiting, so only "would block" means the connection is still good.
fn is_closed(stream: &TcpStream) -> bool {
    if stream.set_nonblocking(true).is_err() {
        return true;
    }
    let mut byte = [0u8; 1];
    let open = matches!(stream.peek(&mut byte), Err(e) if e.kind() == std::io::ErrorKind::WouldBlock);
    stream.set_nonblocking(false).is_err() || !open
}

impl Backend for Remote {
    fn call(&mut self, method: &str, params: Value) -> Result<Value, String> {
        let id = self.next_id;
        self.next_id += 1;
        let line = json!({"id": id, "method": method, "params": params}).to_string();
        // One retry with a fresh connection (the app may have restarted), but only while the request
        // was not sent. Once it is out, a lost reply leaves the outcome unknown and a resend could
        // run an insert/delete/undo twice, so report that instead.
        if self.send(&line).is_err() {
            self.conn = None;
            self.send(&line).map_err(|e| {
                self.conn = None;
                format!("DeckCraft app at {} is not reachable: {e}", self.addr)
            })?;
        }
        let reply = self.receive().map_err(|e| {
            self.conn = None;
            format!(
                "lost the reply from the DeckCraft app at {} ({e}); the command may or may not have run, check the document before retrying",
                self.addr
            )
        })?;
        let v: Value = serde_json::from_str(reply.trim()).map_err(|e| format!("bad reply from app: {e}"))?;
        if v.get("ok").and_then(Value::as_bool) == Some(true) {
            Ok(v.get("result").cloned().unwrap_or(Value::Null))
        } else {
            Err(v.get("error").and_then(Value::as_str).unwrap_or("unknown error").to_string())
        }
    }

    fn has_ui(&self) -> bool {
        true
    }

    fn describe(&self) -> String {
        format!("connected to the DeckCraft app at {}", self.addr)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::net::TcpListener;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicUsize, Ordering};

    #[test]
    fn lost_reply_is_not_resent() {
        let listener = TcpListener::bind("127.0.0.1:0").expect("bind");
        let addr = listener.local_addr().expect("addr").to_string();
        let executed = Arc::new(AtomicUsize::new(0));
        let count = Arc::clone(&executed);
        std::thread::spawn(move || {
            for stream in listener.incoming().flatten() {
                let mut reader = BufReader::new(&stream);
                let mut line = String::new();
                if reader.read_line(&mut line).unwrap_or(0) == 0 {
                    continue;
                }
                let n = count.fetch_add(1, Ordering::SeqCst) + 1;
                if n > 1 {
                    let mut w = &stream;
                    let _ = w.write_all(b"{\"id\":2,\"ok\":true,\"result\":null}\n");
                }
                // The first request is executed, then the connection drops without a reply.
            }
        });
        let mut remote = Remote::connect(&addr).expect("connect");
        let err = remote.call("engine.execute", json!({})).expect_err("lost reply must be an error");
        assert!(err.contains("may or may not have run"), "{err}");
        assert_eq!(executed.load(Ordering::SeqCst), 1, "the command must not be replayed");
        // The next call opens a fresh connection and works.
        assert!(remote.call("document.inspect", json!({})).is_ok());
        assert_eq!(executed.load(Ordering::SeqCst), 2);
    }

    #[test]
    fn connection_closed_before_the_call_is_retried() {
        let listener = TcpListener::bind("127.0.0.1:0").expect("bind");
        let addr = listener.local_addr().expect("addr").to_string();
        let (closed_tx, closed_rx) = std::sync::mpsc::channel();
        let executed = Arc::new(AtomicUsize::new(0));
        let count = Arc::clone(&executed);
        std::thread::spawn(move || {
            let mut first = true;
            for stream in listener.incoming().flatten() {
                if first {
                    // The app restarts: the first connection closes before any request arrives.
                    first = false;
                    drop(stream);
                    let _ = closed_tx.send(());
                    continue;
                }
                let mut reader = BufReader::new(&stream);
                let mut line = String::new();
                if reader.read_line(&mut line).unwrap_or(0) == 0 {
                    continue;
                }
                count.fetch_add(1, Ordering::SeqCst);
                let mut w = &stream;
                let _ = w.write_all(b"{\"id\":1,\"ok\":true,\"result\":null}\n");
            }
        });
        let mut remote = Remote::connect(&addr).expect("connect");
        closed_rx.recv().expect("server closed the first connection");
        std::thread::sleep(Duration::from_millis(200));
        assert!(remote.call("engine.execute", json!({})).is_ok(), "a request that never left must be sent on a fresh connection");
        assert_eq!(executed.load(Ordering::SeqCst), 1);
    }
}
