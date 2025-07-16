use anyhow::{Context, anyhow, bail};
use itertools::Itertools;
use std::{
    io::{ErrorKind, Write},
    path::PathBuf,
    time::Duration,
};
use windows::{
    Data::Xml::Dom::{XmlDocument, XmlElement},
    Foundation::TypedEventHandler,
    Graphics::Imaging::BitmapDecoder,
    Media::Control::{
        GlobalSystemMediaTransportControlsSession, GlobalSystemMediaTransportControlsSessionManager, GlobalSystemMediaTransportControlsSessionMediaProperties,
    },
    Storage::Streams::{Buffer, IBuffer, InputStreamOptions},
    UI::Notifications::{ToastNotification, ToastNotificationManager, ToastTemplateType},
    Win32::System::WinRT::IBufferByteAccess,
    core::Interface,
};

fn create_temp_file_with_extension_and_contents(extension: &str, contents: &[u8]) -> anyhow::Result<PathBuf> {
    let named_temp_file = tempfile::Builder::new()
        .disable_cleanup(true)
        .prefix("thumbnail_")
        .suffix(extension)
        .tempfile()?;
    let path = named_temp_file.path().to_path_buf();
    let mut file = named_temp_file.into_file();
    file.write_all(contents)?;
    Ok(path)
}

fn i_buffer_info_bytes(buffer: &IBuffer) -> anyhow::Result<&[u8]> {
    let i_buffer_byte_access = buffer.cast::<IBufferByteAccess>()?;
    unsafe {
        let data = i_buffer_byte_access.Buffer()?;
        Ok(std::slice::from_raw_parts_mut(data, buffer.Length()? as usize))
    }
}

fn mime_type_to_extension(mime_type: &str) -> anyhow::Result<String> {
    for bitmap_codec_information in BitmapDecoder::GetDecoderInformationEnumerator()? {
        for codec_mime_type in bitmap_codec_information.MimeTypes()? {
            if codec_mime_type.to_os_string() == mime_type {
                return Ok(bitmap_codec_information
                    .FileExtensions()?
                    .into_iter()
                    .next()
                    .ok_or(anyhow!("Mime type found, but has no extensions"))?
                    .to_string_lossy());
            }
        }
    }
    bail!("No extension found")
}

#[derive(PartialEq, Eq, Debug)]
struct Thumbnail {
    mime_type: String,
    bytes: Box<[u8]>,
}

async fn send_toast(
    duration: Duration,
    source_app_user_mode_id: &str,
    line_1: &str,
    line_2: &str,
    line_3: &str,
    thumbnail: &Option<Thumbnail>,
) -> anyhow::Result<()> {
    let toast_template = ToastNotificationManager::GetTemplateContent(if thumbnail.is_some() {
        ToastTemplateType::ToastImageAndText04
    } else {
        ToastTemplateType::ToastText04
    })
    .context("Can not get template content")?;
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
    if let Some(thumbnail) = thumbnail
        && let Ok(extension) = mime_type_to_extension(&thumbnail.mime_type)
    {
        let path = create_temp_file_with_extension_and_contents(&extension, &thumbnail.bytes).context("Can not create temporary file")?;
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
                    .SetAttribute(&"src".into(), &format!("file:///{}", path.as_os_str().to_string_lossy()).into())
                    .context("Can not set attribute `id`")?;
            }
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

#[derive(PartialEq, Eq, Debug)]
struct SessionInfo {
    source_app_user_mode_id: String,
    title: String,
    subtitle: String,
    artist: String,
    album_title: String,
    thumbnail: Option<Thumbnail>,
}
async fn get_thumbnail(
    global_system_media_transport_controls_session_media_properties: &GlobalSystemMediaTransportControlsSessionMediaProperties,
) -> anyhow::Result<Thumbnail> {
    let i_random_access_stream_with_content_type = global_system_media_transport_controls_session_media_properties
        .Thumbnail()?
        .OpenReadAsync()?
        .await?;
    let mime_type = i_random_access_stream_with_content_type.ContentType()?.to_string_lossy();
    let buffer = Buffer::Create(i_random_access_stream_with_content_type.Size()? as _)?;
    i_random_access_stream_with_content_type.ReadAsync(&buffer, i_random_access_stream_with_content_type.Size()? as _, InputStreamOptions::None)?;
    let i_buffer = buffer.into();
    let bytes = i_buffer_info_bytes(&i_buffer)?.to_vec().into_boxed_slice();
    Ok(Thumbnail { mime_type, bytes })
}

async fn get_session_info(global_system_media_transport_controls_session: &GlobalSystemMediaTransportControlsSession) -> anyhow::Result<SessionInfo> {
    let source_app_user_mode_id = global_system_media_transport_controls_session
        .SourceAppUserModelId()
        .context("Can not get source app user model id")?
        .to_string_lossy();
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
    let thumbnail = get_thumbnail(&global_system_media_transport_controls_session_media_properties).await.ok();
    Ok(SessionInfo {
        source_app_user_mode_id,
        title,
        subtitle,
        artist,
        album_title,
        thumbnail,
    })
}

async fn get_session_infos() -> anyhow::Result<Vec<SessionInfo>> {
    let mut session_infos = vec![];
    let global_system_media_transport_controls_session_manager = GlobalSystemMediaTransportControlsSessionManager::RequestAsync()
        .context("Can not get global system media transport controls session manager")?
        .await
        .context("Can not get global system media transport controls session manager")?;
    for global_system_media_transport_controls_session in global_system_media_transport_controls_session_manager
        .GetSessions()
        .context("Can not get sessions")?
    {
        for _ in 0..20 {
            let session_info_result = get_session_info(&global_system_media_transport_controls_session).await;
            match session_info_result {
                Ok(session_info) => {
                    session_infos.push(session_info);
                    break;
                }
                Err(_) => {
                    tokio::time::sleep(Duration::new(0, 50_000_000)).await;
                }
            }
        }
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
    let mut prev_session_infos = vec![];
    while let Some(()) = rx.recv().await {
        let session_infos = get_session_infos().await.context("Can not get session infos")?;
        for session_info in &session_infos {
            if prev_session_infos.contains(session_info) {
                continue;
            }
            if session_info.source_app_user_mode_id.starts_with("MSTeams_") {
                continue;
            }
            let full_title = format!("{} – {}", &session_info.title, &session_info.subtitle);
            send_toast(
                Duration::new(4, 0),
                &session_info.source_app_user_mode_id,
                if session_info.subtitle.is_empty() { &session_info.title } else { &full_title },
                &session_info.album_title,
                &session_info.artist,
                &session_info.thumbnail,
            )
            .await
            .context("Failed to send toast")?;
        }
        prev_session_infos = session_infos;
    }
    Ok(())
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    run_notifer().await.context("Notifier failed")?;
    Ok(())
}
