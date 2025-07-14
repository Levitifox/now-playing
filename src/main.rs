use itertools::Itertools;
use std::time::Duration;
use windows::{
    Data::Xml::Dom::{XmlDocument, XmlElement},
    Foundation::TypedEventHandler,
    Media::Control::GlobalSystemMediaTransportControlsSessionManager,
    UI::Notifications::{ToastNotification, ToastNotificationManager, ToastTemplateType},
    Win32::Foundation::{ERROR_BROKEN_PIPE, ERROR_NOT_FOUND},
    core::{HRESULT, Interface},
};

async fn send_toast(duration: Duration, line_1: &str, line_2: &str, line_3: &str) -> anyhow::Result<()> {
    let toast_template = ToastNotificationManager::GetTemplateContent(ToastTemplateType::ToastImageAndText04)?;
    let toast_element = toast_template
        .GetElementsByTagName(&"toast".into())?
        .into_iter()
        .exactly_one()
        .map_err(|_| windows_result::Error::from_hresult(HRESULT::from_win32(ERROR_NOT_FOUND.0)))?
        .cast::<XmlElement>()?;
    for text_node in toast_element.GetElementsByTagName(&"text".into())?.into_iter().collect::<Vec<_>>() {
        let text_element = text_node.cast::<XmlElement>()?;
        if text_element.GetAttribute(&"id".into())?.to_string_lossy() == "1" {
            text_element.AppendChild(&XmlDocument::CreateTextNode(&toast_template, &line_1.into())?)?;
        }
        if text_element.GetAttribute(&"id".into())?.to_string_lossy() == "2" {
            text_element.AppendChild(&XmlDocument::CreateTextNode(&toast_template, &line_2.into())?)?;
        }
        if text_element.GetAttribute(&"id".into())?.to_string_lossy() == "3" {
            text_element.AppendChild(&XmlDocument::CreateTextNode(&toast_template, &line_3.into())?)?;
        }
    }
    for image_node in toast_element.GetElementsByTagName(&"image".into())?.into_iter().collect::<Vec<_>>() {
        let image_element = image_node.cast::<XmlElement>()?;
        if image_element.GetAttribute(&"id".into())?.to_string_lossy() == "1" {
            image_element.SetAttribute(&"src".into(), &"file:///C:/Users/artfd/Pictures/Unpacking/20250711_0001.png".into())?;
        }
    }
    let audio_element = toast_template.CreateElement(&"audio".into())?;
    audio_element.SetAttribute(&"silent".into(), &"true".into())?;
    toast_element.AppendChild(&audio_element)?;
    let toast_notifier = ToastNotificationManager::CreateToastNotifierWithId(&"now-playing".into())?;
    let toast_notification = ToastNotification::CreateToastNotification(&toast_template)?;
    toast_notifier.Show(&toast_notification)?;
    tokio::time::sleep(duration).await;
    toast_notifier.Hide(&toast_notification)?;
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
    let global_system_media_transport_controls_session_manager = GlobalSystemMediaTransportControlsSessionManager::RequestAsync()?.await?;
    for global_system_media_transport_controls_session in global_system_media_transport_controls_session_manager.GetSessions()? {
        let source_app_user_mode_id = global_system_media_transport_controls_session.SourceAppUserModelId()?.to_string_lossy();
        let global_system_media_transport_controls_session_media_properties =
            global_system_media_transport_controls_session.TryGetMediaPropertiesAsync()?.await?;
        let title = global_system_media_transport_controls_session_media_properties.Title()?.to_string_lossy();
        let subtitle = global_system_media_transport_controls_session_media_properties.Subtitle()?.to_string_lossy();
        let artist = global_system_media_transport_controls_session_media_properties.Artist()?.to_string_lossy();
        let album_title = global_system_media_transport_controls_session_media_properties.AlbumTitle()?.to_string_lossy();
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
    let global_system_media_transport_controls_session_manager = GlobalSystemMediaTransportControlsSessionManager::RequestAsync()?.await?;
    global_system_media_transport_controls_session_manager.SessionsChanged(&TypedEventHandler::new({
        let tx = tx.clone();
        move |_, _| {
            tx.send(())
                .map_err(|_| windows_result::Error::from_hresult(HRESULT::from_win32(ERROR_BROKEN_PIPE.0)))?;
            Ok(())
        }
    }))?;
    tx.send(())
        .map_err(|_| windows_result::Error::from_hresult(HRESULT::from_win32(ERROR_BROKEN_PIPE.0)))?;
    while let Some(()) = rx.recv().await {
        let session_infos = sessions_changed().await?;
        for session_info in session_infos {
            dbg!(&session_info.source_app_user_mode_id);
            if session_info.source_app_user_mode_id.starts_with("MSTeams_") {
                continue;
            }
            send_toast(
                Duration::new(5, 0),
                &(if session_info.subtitle.is_empty() {
                    session_info.title
                } else {
                    format!("{} – {}", session_info.title, session_info.subtitle)
                }),
                &session_info.album_title,
                &session_info.artist,
            )
            .await?;
        }
    }
    Ok(())
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    run_notifer().await?;
    Ok(())
}
