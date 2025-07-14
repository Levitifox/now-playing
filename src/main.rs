use anyhow::{Context, anyhow};
use itertools::Itertools;
use std::{io::ErrorKind, time::Duration};
use windows::{
    Data::Xml::Dom::{XmlDocument, XmlElement},
    Foundation::TypedEventHandler,
    Media::Control::GlobalSystemMediaTransportControlsSessionManager,
    UI::Notifications::{ToastNotification, ToastNotificationManager, ToastTemplateType},
    core::Interface,
};

async fn send_toast(duration: Duration, source_app_user_mode_id: &str, line_1: &str, line_2: &str, line_3: &str) -> anyhow::Result<()> {
    let toast_template = ToastNotificationManager::GetTemplateContent(ToastTemplateType::ToastImageAndText04).context("Can not get template content")?;
    let toast_element = toast_template
        .GetElementsByTagName(&"toast".into())
        .context("Can not find element <toast>")?
        .into_iter()
        .exactly_one()
        .map_err(|_| anyhow!("Not exactly one element <toast>"))?
        .cast::<XmlElement>()
        .context("Node <toast> is not an element")?;
    for text_node in toast_element
        .GetElementsByTagName(&"text".into())
        .context("Can not find elements <text>")?
        .into_iter()
        .collect::<Vec<_>>()
    {
        let text_element = text_node.cast::<XmlElement>().context("Node <text> is not an element")?;
        if text_element.GetAttribute(&"id".into()).context("Can not get attribute `id`")?.to_string_lossy() == "1" {
            text_element
                .AppendChild(&XmlDocument::CreateTextNode(&toast_template, &line_1.into()).context("Can not create text node")?)
                .context("Can not append child")?;
        }
        if text_element.GetAttribute(&"id".into()).context("Can not get attribute `id`")?.to_string_lossy() == "2" {
            text_element
                .AppendChild(&XmlDocument::CreateTextNode(&toast_template, &line_2.into()).context("Can not create text node")?)
                .context("Can not append child")?;
        }
        if text_element.GetAttribute(&"id".into()).context("Can not get attribute `id`")?.to_string_lossy() == "3" {
            text_element
                .AppendChild(&XmlDocument::CreateTextNode(&toast_template, &line_3.into()).context("Can not create text node")?)
                .context("Can not append child")?;
        }
    }
    for image_node in toast_element
        .GetElementsByTagName(&"image".into())
        .context("Can not find elements <image>")?
        .into_iter()
        .collect::<Vec<_>>()
    {
        let image_element = image_node.cast::<XmlElement>().context("Node <image> is not an element")?;
        if image_element
            .GetAttribute(&"id".into())
            .context("Can not get attribute `id`")?
            .to_string_lossy()
            == "1"
        {
            image_element
                .SetAttribute(&"src".into(), &"file:///C:/Users/artfd/Pictures/Unpacking/20250711_0001.png".into())
                .context("Can not set attribute `id`")?;
        }
    }
    let audio_element = toast_template.CreateElement(&"audio".into()).context("Can not create element <audio>")?;
    audio_element
        .SetAttribute(&"silent".into(), &"true".into())
        .context("Can not set attribute `silent`")?;
    toast_element.AppendChild(&audio_element).context("Can not append child")?;
    let toast_notifier = ToastNotificationManager::CreateToastNotifierWithId(&source_app_user_mode_id.into()).context("Can not creat toast notifier")?;
    let toast_notification = ToastNotification::CreateToastNotification(&toast_template).context("Can not creat toast notification")?;
    toast_notifier.Show(&toast_notification).context("Can not show notification")?;
    tokio::time::sleep(duration).await;
    toast_notifier.Hide(&toast_notification).context("Can not hide notification")?;
    Ok(())
}

#[derive(Debug)]
struct SessionInfo {
    source_app_user_mode_id: String,
    title: String,
    subtitle: String,
    artist: String,
    album_title: String,
}

async fn sessions_changed() -> anyhow::Result<Vec<SessionInfo>> {
    let mut session_infos = vec![];
    let global_system_media_transport_controls_session_manager = GlobalSystemMediaTransportControlsSessionManager::RequestAsync()
        .context("Can not get global system media transport controls session manager")?
        .await
        .context("Can not get global system media transport controls session manager")?;
    for global_system_media_transport_controls_session in global_system_media_transport_controls_session_manager
        .GetSessions()
        .context("Can not get sessions")?
    {
        let source_app_user_mode_id = global_system_media_transport_controls_session
            .SourceAppUserModelId()
            .context("Can not get source app user model id")?
            .to_string_lossy();
        tokio::time::sleep(Duration::new(2, 0)).await;
        let global_system_media_transport_controls_session_media_properties = global_system_media_transport_controls_session
            .TryGetMediaPropertiesAsync()
            .context("Can not get media properties")?
            .await
            .context("Can not get media properties")?;
        let title = global_system_media_transport_controls_session_media_properties
            .Title()
            .context("Can not get title")?
            .to_string_lossy();
        let subtitle = global_system_media_transport_controls_session_media_properties
            .Subtitle()
            .context("Can not get subtitle")?
            .to_string_lossy();
        let artist = global_system_media_transport_controls_session_media_properties
            .Artist()
            .context("Can not get artist")?
            .to_string_lossy();
        let album_title = global_system_media_transport_controls_session_media_properties
            .AlbumTitle()
            .context("Can not get album title")?
            .to_string_lossy();
        session_infos.push(SessionInfo {
            source_app_user_mode_id,
            title,
            subtitle,
            artist,
            album_title,
        });
    }
    Ok(session_infos)
}

async fn run_notifer() -> anyhow::Result<()> {
    let (tx, mut rx) = tokio::sync::mpsc::unbounded_channel::<()>();
    let global_system_media_transport_controls_session_manager = GlobalSystemMediaTransportControlsSessionManager::RequestAsync()
        .context("Can not get global system media transport controls session manager")?
        .await
        .context("Can not get global system media transport controls session manager")?;
    global_system_media_transport_controls_session_manager.SessionsChanged(&TypedEventHandler::new({
        let tx = tx.clone();
        move |_, _| {
            tx.send(())
                .map_err(|e| windows_result::Error::from(std::io::Error::new(ErrorKind::BrokenPipe, e)))?;
            Ok(())
        }
    }))?;
    tx.send(())?;
    while let Some(()) = rx.recv().await {
        let session_infos = sessions_changed().await.context("Can not get session infos")?;
        for session_info in session_infos {
            if session_info.source_app_user_mode_id.starts_with("MSTeams_") {
                continue;
            }
            send_toast(
                Duration::new(5, 0),
                &session_info.source_app_user_mode_id,
                &(if session_info.subtitle.is_empty() {
                    session_info.title
                } else {
                    format!("{} – {}", session_info.title, session_info.subtitle)
                }),
                &session_info.album_title,
                &session_info.artist,
            )
            .await
            .context("Failed to send toast")?;
        }
    }
    Ok(())
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    run_notifer().await.context("Notifier failed")?;
    Ok(())
}
