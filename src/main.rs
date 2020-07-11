use std::io::{Cursor, Write};

use async_std::net::UdpSocket;

pub mod tts;

#[async_std::main]
async fn main() {
    let mut args = std::env::args().skip(1);
    let result = match args.next().as_deref() {
        Some("--help") | None => {
            eprintln!("USAGE: boop [name]            broadcast a boop packet to [name]");
            eprintln!("       boop --everyone        boop everyone (don't be annoying tho!)");
            eprintln!("       boop --help            get this message");
            eprintln!("       boop --daemon [name]   start the boop daemon, listening as [name]");
            eprintln!("       boop --install [name]  adds the boop daemon to startup as [name]");
            Ok(())
        }

        Some("--daemon") => match args.next() {
            Some(name) => daemon(name).await,
            None => {
                eprintln!("Daemon requires a listening name");
                Ok(())
            }
        },

        Some("--install") => match args.next() {
            Some(name) => install(name),
            None => {
                eprintln!("Installation requires a listening name");
                Ok(())
            }
        },

        Some("--everyone") => boop(None).await,
        Some(name) => boop(Some(name.to_string())).await,
    };

    if let Err(err) = result {
        eprintln!("Failure in boop command: {}", err);
    }
}

const PORT: u16 = 0xC_C_24; // C_C_Z

async fn get_my_name(sock: &UdpSocket) -> anyhow::Result<String> {
    sock.send_to(b"get-name", ("127.0.0.1", PORT)).await?;

    let mut buf = [0; 512];
    let len = sock.recv(&mut buf).await?;

    Ok(String::from_utf8_lossy(&buf[..len]).into_owned())
}

async fn boop(name: Option<String>) -> anyhow::Result<()> {
    let sock = UdpSocket::bind(("0.0.0.0", 0)).await?;
    let my_name = get_my_name(&sock).await?;
    sock.set_broadcast(true)?;

    let mut buf = [0; 512];
    let buflen = {
        let mut cursor = Cursor::new(&mut buf[..]);
        if let Some(name) = name {
            write!(&mut cursor, "boop {}->{}", my_name, name)?;
        } else {
            write!(&mut cursor, "boop {}", my_name)?;
        }
        cursor.position() as usize
    };

    sock.send_to(&buf[..buflen], ("255.255.255.255", PORT))
        .await?;
    Ok(())
}

async fn daemon(name: String) -> anyhow::Result<()> {
    let mut buf = [0; 512];

    let sock = UdpSocket::bind(("0.0.0.0", PORT)).await?;
    sock.set_broadcast(true)?;

    let mut speaker = tts::Speaker::new()?;

    loop {
        let (len, sender) = match sock.recv_from(&mut buf).await {
            Err(err) => break Err(err.into()),
            Ok((len, _)) if len < 6 => continue,
            Ok(res) => res,
        };

        let cmd = match std::str::from_utf8(&buf[..len]) {
            Ok(cmd) => cmd,
            Err(_) => continue,
        };

        if cmd == "get-name" {
            sock.send_to(name.as_bytes(), sender).await?;
        } else if cmd.starts_with("boop ") {
            let mut parts = cmd[5..].split("->");
            let sender_name = parts.next().expect("Split always has at least 1 component");
            let recipient = parts.next();

            if recipient.is_some() && recipient != Some(name.as_ref()) {
                continue;
            }

            let message = if recipient.is_none() {
                format!("{} is booping everyone!!!", sender_name)
            } else {
                format!("{} is booping you!", sender_name)
            };

            speaker.speak_async(&message)?;
        }
    }
}

fn install(name: String) -> anyhow::Result<()> {
    let posh = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.EXE";
    let exe = std::env::current_exe()?;
    let cmd = format!(
        ". '{path}' --daemon '{name}'",
        path = exe.to_string_lossy(),
        name = name
    );
    let args = format!(
        "-windowstyle hidden -command \"{cmd}\"",
        cmd = cmd
    );

    let appdata = std::env::var("APPDATA")?;
    let startup = format!(
        "{appdata}\\Microsoft\\Windows\\Start Menu\\Programs\\Startup",
        appdata = appdata
    );
    let shortcut = format!("{startup}\\Boop.lnk", startup = startup);

    let make_shortcut = format!(
        "$ws = New-Object -ComObject WScript.Shell; \
         $shortcut = $ws.CreateShortcut(\"{shortcut}\"); \
         $shortcut.TargetPath = \"{posh}\"; \
         $shortcut.WorkingDirectory = \"{exedir}\"; \
         $shortcut.Arguments = \"{args}\"; \
         $shortcut.Save();",
        shortcut = shortcut,
        posh = posh,
        exedir = exe.parent().unwrap().to_string_lossy(),
        args = args.replace("\"", "`\""),
    );

    std::process::Command::new(posh)
        .arg("-Command")
        .arg(&make_shortcut)
        .spawn()?
        .wait()?;

    std::process::Command::new("cmd.exe")
        .arg("/C")
        .arg(&shortcut)
        .spawn()?;

    Ok(())
}
