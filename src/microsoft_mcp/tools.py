import base64
import datetime as dt
import html as html_mod
import pathlib as pl
import re
from typing import Any
from fastmcp import FastMCP
from . import graph, auth

mcp = FastMCP("microsoft-mcp")

_HTML_TAG_RE = re.compile(r"<(html|body|p|br|div|span|ul|ol|li|table|h[1-6]|b|i|u|a|strong|em|hr|pre|code|blockquote)\b", re.IGNORECASE)

_SIGNATURE_LOGO_PATH = pl.Path("/Users/paul/Documents/Sync/Brand/variablegrid-logo.png")
_SIGNATURE_LOGO_CID = "variablegrid-logo"

_SIGNATURE_HTML = f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #1f1f1f;">
<p style="margin:0 0 6pt 0;">Warm Regards,</p>
<p style="margin:0;"><b>Paul McLaughlin</b></p>
<p style="margin:6pt 0;"><img src="cid:{_SIGNATURE_LOGO_CID}" alt="Variable Grid" style="height:52px;"></p>
<p style="margin:0;"><b>Variablegrid Adaptive Power Inc.</b></p>
<p style="margin:0;">Email: <a href="mailto:pmclaughlin@variablegrid.com">pmclaughlin@variablegrid.com</a></p>
</div>"""

_SIGNATURE_PLAIN = """Warm Regards,

Paul McLaughlin

Variablegrid Adaptive Power Inc.
Email: pmclaughlin@variablegrid.com"""


def _email_body(content: str) -> dict[str, str]:
    """Build a Graph message body, auto-detecting HTML vs plain text."""
    content_type = "HTML" if _HTML_TAG_RE.search(content) else "Text"
    return {"contentType": content_type, "content": content}


def _chat_body(content: str) -> dict[str, str]:
    """Build a Teams chatMessage body.

    Teams renders contentType=text messages with newlines collapsed, so
    plain text is escaped and converted to HTML with <br> line breaks.
    Content that already contains HTML tags is passed through as HTML.

    NOTE: chatMessage requires lowercase "html"/"text" — uppercase "HTML"
    (fine for email) gets a 400 from /chats/{id}/messages.
    """
    if _HTML_TAG_RE.search(content):
        return {"contentType": "html", "content": content}
    escaped = html_mod.escape(content).replace("\n", "<br>")
    return {"contentType": "html", "content": escaped}


def _strip_existing_signoff(body: str, is_html: bool) -> str:
    """Remove a trailing 'Best,/Warm Regards,/Thanks,/Paul' block so we don't double-sign."""
    if is_html:
        # Strip trailing <p>...sign-off...</p> blocks containing Paul + a closer
        pattern = re.compile(
            r"(<p[^>]*>\s*(Warm Regards|Best|Thanks|Cheers)[,\s]*<br\s*/?>\s*Paul[\s\S]*?</p>\s*)+\s*$",
            re.IGNORECASE,
        )
        return pattern.sub("", body).rstrip()
    pattern = re.compile(
        r"\n\s*(Warm Regards|Best|Thanks|Cheers)\s*,\s*\n\s*Paul\s*$",
        re.IGNORECASE,
    )
    return pattern.sub("", body).rstrip()


def _apply_signature(body: str) -> tuple[str, list[dict[str, Any]]]:
    """Append Paul's signature to body. Returns (new_body, inline_attachments).

    For HTML bodies, embeds the logo as an inline cid: image and returns the
    Graph file-attachment dict so the caller can attach it to the message.
    For plain-text bodies, appends the plain-text signature with no logo.
    """
    is_html = bool(_HTML_TAG_RE.search(body))
    body = _strip_existing_signoff(body, is_html)

    if not is_html:
        return f"{body}\n\n{_SIGNATURE_PLAIN}", []

    # Insert signature before </body> if present, otherwise append
    if re.search(r"</body>\s*</html>\s*$", body, re.IGNORECASE):
        new_body = re.sub(
            r"(</body>\s*</html>\s*)$",
            f"{_SIGNATURE_HTML}\\1",
            body,
            count=1,
            flags=re.IGNORECASE,
        )
    elif re.search(r"</body>\s*$", body, re.IGNORECASE):
        new_body = re.sub(
            r"(</body>\s*)$", f"{_SIGNATURE_HTML}\\1", body, count=1, flags=re.IGNORECASE
        )
    else:
        new_body = f"{body}\n{_SIGNATURE_HTML}"

    inline_attachments: list[dict[str, Any]] = []
    if _SIGNATURE_LOGO_PATH.exists():
        logo_bytes = _SIGNATURE_LOGO_PATH.read_bytes()
        inline_attachments.append(
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": _SIGNATURE_LOGO_PATH.name,
                "contentType": "image/png",
                "contentBytes": base64.b64encode(logo_bytes).decode("utf-8"),
                "contentId": _SIGNATURE_LOGO_CID,
                "isInline": True,
            }
        )
    return new_body, inline_attachments


def _file_attachment(
    file_path: str, content_id: str | None = None, inline: bool = False
) -> dict[str, Any]:
    """Build a Graph fileAttachment dict from a path (optionally inline w/ cid)."""
    path = pl.Path(file_path).expanduser().resolve()
    att: dict[str, Any] = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": path.name,
        "contentBytes": base64.b64encode(path.read_bytes()).decode("utf-8"),
    }
    if content_id:
        att["contentId"] = content_id
    if inline:
        att["isInline"] = True
    return att


def _compose_reply_body(reply_body: str, quoted: str, source_is_html: bool) -> str:
    """Combine new reply text with the quoted history a createReply produced.

    For HTML threads, insert the reply *inside* the quoted <body> so the markup
    stays valid: prepending fragments before a full <html> document yields
    malformed HTML that a sanitizer can reflow, stripping inline cid: images.
    For plain-text threads, just join with blank lines.
    """
    if source_is_html:
        m = re.search(r"<body[^>]*>", quoted, re.IGNORECASE)
        if m:
            i = m.end()
            return quoted[:i] + reply_body + "<br><br>" + quoted[i:]
        return reply_body + "<br><br>" + quoted
    return reply_body + "\n\n" + quoted


FOLDERS = {
    k.casefold(): v
    for k, v in {
        "inbox": "inbox",
        "sent": "sentitems",
        "drafts": "drafts",
        "deleted": "deleteditems",
        "junk": "junkemail",
        "archive": "archive",
    }.items()
}


@mcp.tool
def list_accounts() -> list[dict[str, str]]:
    """List all signed-in Microsoft accounts"""
    return [
        {"username": acc.username, "account_id": acc.account_id}
        for acc in auth.list_accounts()
    ]


@mcp.tool
def authenticate_account() -> dict[str, str]:
    """Authenticate a new Microsoft account using device flow authentication

    Returns authentication instructions and device code for the user to complete authentication.
    The user must visit the URL and enter the code to authenticate their Microsoft account.
    """
    app = auth.get_app()
    flow = app.initiate_device_flow(scopes=auth.SCOPES)

    if "user_code" not in flow:
        error_msg = flow.get("error_description", "Unknown error")
        raise Exception(f"Failed to get device code: {error_msg}")

    verification_url = flow.get(
        "verification_uri",
        flow.get("verification_url", "https://microsoft.com/devicelogin"),
    )

    return {
        "status": "authentication_required",
        "instructions": "To authenticate a new Microsoft account:",
        "step1": f"Visit: {verification_url}",
        "step2": f"Enter code: {flow['user_code']}",
        "step3": "Sign in with the Microsoft account you want to add",
        "step4": "After authenticating, use the 'complete_authentication' tool to finish the process",
        "device_code": flow["user_code"],
        "verification_url": verification_url,
        "expires_in": flow.get("expires_in", 900),
        "_flow_cache": str(flow),
    }


@mcp.tool
def complete_authentication(flow_cache: str) -> dict[str, str]:
    """Complete the authentication process after the user has entered the device code

    Args:
        flow_cache: The flow data returned from authenticate_account (the _flow_cache field)

    Returns:
        Account information if authentication was successful
    """
    import ast

    try:
        flow = ast.literal_eval(flow_cache)
    except (ValueError, SyntaxError):
        raise ValueError("Invalid flow cache data")

    app = auth.get_app()
    result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        error_msg = result.get("error_description", result["error"])
        if "authorization_pending" in error_msg:
            return {
                "status": "pending",
                "message": "Authentication is still pending. The user needs to complete the authentication process.",
                "instructions": "Please ensure you've visited the URL and entered the code, then try again.",
            }
        raise Exception(f"Authentication failed: {error_msg}")

    # Save the token cache
    cache = app.token_cache
    if isinstance(cache, auth.msal.SerializableTokenCache) and cache.has_state_changed:
        auth._write_cache(cache.serialize())

    # Get the newly added account
    accounts = app.get_accounts()
    if accounts:
        # Find the account that matches the token we just got
        for account in accounts:
            if (
                account.get("username", "").lower()
                == result.get("id_token_claims", {})
                .get("preferred_username", "")
                .lower()
            ):
                return {
                    "status": "success",
                    "username": account["username"],
                    "account_id": account["home_account_id"],
                    "message": f"Successfully authenticated {account['username']}",
                }
        # If exact match not found, return the last account
        account = accounts[-1]
        return {
            "status": "success",
            "username": account["username"],
            "account_id": account["home_account_id"],
            "message": f"Successfully authenticated {account['username']}",
        }

    return {
        "status": "error",
        "message": "Authentication succeeded but no account was found",
    }


@mcp.tool
def list_emails(
    account_id: str,
    folder: str = "inbox",
    limit: int = 10,
    include_body: bool = True,
) -> list[dict[str, Any]]:
    """List emails from specified folder"""
    folder_path = FOLDERS.get(folder.casefold(), folder)

    if include_body:
        select_fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,body,conversationId,isRead"
    else:
        select_fields = "id,subject,from,toRecipients,receivedDateTime,hasAttachments,conversationId,isRead"

    params = {
        "$top": min(limit, 100),
        "$select": select_fields,
        "$orderby": "receivedDateTime desc",
    }

    emails = list(
        graph.request_paginated(
            f"/me/mailFolders/{folder_path}/messages",
            account_id,
            params=params,
            limit=limit,
        )
    )

    return emails


@mcp.tool
def get_email(
    email_id: str,
    account_id: str,
    include_body: bool = True,
    body_max_length: int = 50000,
    include_attachments: bool = True,
) -> dict[str, Any]:
    """Get email details with size limits

    Args:
        email_id: The email ID
        account_id: The account ID
        include_body: Whether to include the email body (default: True)
        body_max_length: Maximum characters for body content (default: 50000)
        include_attachments: Whether to include attachment metadata (default: True)
    """
    params = {}
    if include_attachments:
        params["$expand"] = "attachments($select=id,name,size,contentType)"

    result = graph.request("GET", f"/me/messages/{email_id}", account_id, params=params)
    if not result:
        raise ValueError(f"Email with ID {email_id} not found")

    # Truncate body if needed
    if include_body and "body" in result and "content" in result["body"]:
        content = result["body"]["content"]
        if len(content) > body_max_length:
            result["body"]["content"] = (
                content[:body_max_length]
                + f"\n\n[Content truncated - {len(content)} total characters]"
            )
            result["body"]["truncated"] = True
            result["body"]["total_length"] = len(content)
    elif not include_body and "body" in result:
        del result["body"]

    # Remove attachment content bytes to reduce size
    if "attachments" in result and result["attachments"]:
        for attachment in result["attachments"]:
            if "contentBytes" in attachment:
                del attachment["contentBytes"]

    return result


@mcp.tool
def create_email_draft(
    account_id: str,
    to: str | list[str],
    subject: str,
    body: str,
    cc: str | list[str] | None = None,
    attachments: str | list[str] | None = None,
    append_signature: bool = True,
) -> dict[str, Any]:
    """Create an email draft with file path(s) as attachments.

    append_signature (default True): appends Paul's "Warm Regards" sign-off,
    embedding the variablegrid logo as an inline image when the body is HTML.
    """
    to_list = [to] if isinstance(to, str) else to

    inline_attachments: list[dict[str, Any]] = []
    if append_signature:
        body, inline_attachments = _apply_signature(body)

    message = {
        "subject": subject,
        "body": _email_body(body),
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_list],
    }

    if cc:
        cc_list = [cc] if isinstance(cc, str) else cc
        message["ccRecipients"] = [
            {"emailAddress": {"address": addr}} for addr in cc_list
        ]

    small_attachments = list(inline_attachments)
    large_attachments = []

    if attachments:
        # Convert single path to list
        attachment_paths = (
            [attachments] if isinstance(attachments, str) else attachments
        )
        for file_path in attachment_paths:
            path = pl.Path(file_path).expanduser().resolve()
            content_bytes = path.read_bytes()
            att_size = len(content_bytes)
            att_name = path.name

            if att_size < 3 * 1024 * 1024:
                small_attachments.append(
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": att_name,
                        "contentBytes": base64.b64encode(content_bytes).decode("utf-8"),
                    }
                )
            else:
                large_attachments.append(
                    {
                        "name": att_name,
                        "content_bytes": content_bytes,
                        "content_type": "application/octet-stream",
                    }
                )

    if small_attachments:
        message["attachments"] = small_attachments

    result = graph.request("POST", "/me/messages", account_id, json=message)
    if not result:
        raise ValueError("Failed to create email draft")

    message_id = result["id"]

    for att in large_attachments:
        graph.upload_large_mail_attachment(
            message_id,
            att["name"],
            att["content_bytes"],
            account_id,
            att.get("content_type", "application/octet-stream"),
        )

    return result


@mcp.tool
def send_email(
    account_id: str,
    to: str | list[str],
    subject: str,
    body: str,
    cc: str | list[str] | None = None,
    attachments: str | list[str] | None = None,
    append_signature: bool = True,
) -> dict[str, str]:
    """Send an email immediately with file path(s) as attachments.

    append_signature (default True): appends Paul's "Warm Regards" sign-off,
    embedding the variablegrid logo as an inline image when the body is HTML.
    """
    to_list = [to] if isinstance(to, str) else to

    inline_attachments: list[dict[str, Any]] = []
    if append_signature:
        body, inline_attachments = _apply_signature(body)

    message = {
        "subject": subject,
        "body": _email_body(body),
        "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_list],
    }

    if cc:
        cc_list = [cc] if isinstance(cc, str) else cc
        message["ccRecipients"] = [
            {"emailAddress": {"address": addr}} for addr in cc_list
        ]

    # Check if we have large attachments
    has_large_attachments = False
    processed_attachments = []

    if attachments:
        # Convert single path to list
        attachment_paths = (
            [attachments] if isinstance(attachments, str) else attachments
        )
        for file_path in attachment_paths:
            path = pl.Path(file_path).expanduser().resolve()
            content_bytes = path.read_bytes()
            att_size = len(content_bytes)
            att_name = path.name

            processed_attachments.append(
                {
                    "name": att_name,
                    "content_bytes": content_bytes,
                    "content_type": "application/octet-stream",
                    "size": att_size,
                }
            )

            if att_size >= 3 * 1024 * 1024:
                has_large_attachments = True

    if not has_large_attachments and processed_attachments:
        message["attachments"] = list(inline_attachments) + [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": att["name"],
                "contentBytes": base64.b64encode(att["content_bytes"]).decode("utf-8"),
            }
            for att in processed_attachments
        ]
        graph.request("POST", "/me/sendMail", account_id, json={"message": message})
        return {"status": "sent"}
    elif has_large_attachments:
        # Create draft first, then add large attachments, then send
        # We need to handle large attachments manually here
        to_list = [to] if isinstance(to, str) else to
        message = {
            "subject": subject,
            "body": _email_body(body),
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to_list],
        }
        if cc:
            cc_list = [cc] if isinstance(cc, str) else cc
            message["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc_list
            ]
        if inline_attachments:
            message["attachments"] = list(inline_attachments)

        result = graph.request("POST", "/me/messages", account_id, json=message)
        if not result:
            raise ValueError("Failed to create email draft")

        message_id = result["id"]

        for att in processed_attachments:
            if att["size"] >= 3 * 1024 * 1024:
                graph.upload_large_mail_attachment(
                    message_id,
                    att["name"],
                    att["content_bytes"],
                    account_id,
                    att.get("content_type", "application/octet-stream"),
                )
            else:
                small_att = {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": att["name"],
                    "contentBytes": base64.b64encode(att["content_bytes"]).decode(
                        "utf-8"
                    ),
                }
                graph.request(
                    "POST",
                    f"/me/messages/{message_id}/attachments",
                    account_id,
                    json=small_att,
                )

        graph.request("POST", f"/me/messages/{message_id}/send", account_id)
        return {"status": "sent"}
    else:
        if inline_attachments:
            message["attachments"] = list(inline_attachments)
        graph.request("POST", "/me/sendMail", account_id, json={"message": message})
        return {"status": "sent"}


@mcp.tool
def update_email(
    email_id: str, updates: dict[str, Any], account_id: str
) -> dict[str, Any]:
    """Update email properties (isRead, categories, flag, etc.)"""
    result = graph.request(
        "PATCH", f"/me/messages/{email_id}", account_id, json=updates
    )
    if not result:
        raise ValueError(f"Failed to update email {email_id} - no response")
    return result


@mcp.tool
def delete_email(email_id: str, account_id: str) -> dict[str, str]:
    """Delete an email"""
    graph.request("DELETE", f"/me/messages/{email_id}", account_id)
    return {"status": "deleted"}


@mcp.tool
def move_email(
    email_id: str, destination_folder: str, account_id: str
) -> dict[str, Any]:
    """Move email to another folder"""
    folder_path = FOLDERS.get(destination_folder.casefold(), destination_folder)

    folders = graph.request("GET", "/me/mailFolders", account_id)
    folder_id = None

    if not folders:
        raise ValueError("Failed to retrieve mail folders")
    if "value" not in folders:
        raise ValueError(f"Unexpected folder response structure: {folders}")

    for folder in folders["value"]:
        if folder["displayName"].lower() == folder_path.lower():
            folder_id = folder["id"]
            break

    if not folder_id:
        raise ValueError(f"Folder '{destination_folder}' not found")

    payload = {"destinationId": folder_id}
    result = graph.request(
        "POST", f"/me/messages/{email_id}/move", account_id, json=payload
    )
    if not result:
        raise ValueError("Failed to move email - no response from server")
    if "id" not in result:
        raise ValueError(f"Failed to move email - unexpected response: {result}")
    return {"status": "moved", "new_id": result["id"]}


@mcp.tool
def reply_to_email(account_id: str, email_id: str, body: str) -> dict[str, str]:
    """Reply to an email (sender only)"""
    endpoint = f"/me/messages/{email_id}/reply"
    payload = {"message": {"body": _email_body(body)}}
    graph.request("POST", endpoint, account_id, json=payload)
    return {"status": "sent"}


@mcp.tool
def reply_all_email(account_id: str, email_id: str, body: str) -> dict[str, str]:
    """Reply to all recipients of an email"""
    endpoint = f"/me/messages/{email_id}/replyAll"
    payload = {"message": {"body": _email_body(body)}}
    graph.request("POST", endpoint, account_id, json=payload)
    return {"status": "sent"}


@mcp.tool
def create_reply_draft(
    account_id: str,
    email_id: str,
    body: str,
    reply_all: bool = False,
    attachments: str | list[str] | None = None,
    inline_images: str | list[str] | None = None,
    append_signature: bool = True,
) -> dict[str, Any]:
    """Create a DRAFT reply to an existing email, threaded into the conversation.

    Unlike reply_to_email / reply_all_email (which send immediately with no
    attachments), this leaves a draft for review and supports file attachments
    and inline images. The quoted thread history is preserved.

    inline_images: file path(s) embedded inline in the body. Reference each in
    the HTML body as <img src="cid:NAME"> where NAME is the image file's stem
    (an image at /tmp/panel.png is referenced as <img src="cid:panel">).

    append_signature (default True): appends Paul's sign-off (+ logo for HTML).

    Note: the reply inherits the original message's content type. Replying to a
    plain-text message yields a plain-text draft, so HTML formatting and inline
    images will not render; reply to an HTML message for rich formatting.
    """
    endpoint = (
        f"/me/messages/{email_id}/{'createReplyAll' if reply_all else 'createReply'}"
    )
    draft = graph.request("POST", endpoint, account_id)
    if not draft:
        raise ValueError("Failed to create reply draft")
    draft_id = draft["id"]
    quoted = draft.get("body", {}).get("content", "")
    source_is_html = (
        draft.get("body", {}).get("contentType", "html").lower() == "html"
    )

    inline_atts: list[dict[str, Any]] = []
    if append_signature:
        body, inline_atts = _apply_signature(body)
    full_body = _compose_reply_body(body, quoted, source_is_html)

    if inline_images:
        paths = [inline_images] if isinstance(inline_images, str) else inline_images
        for p in paths:
            inline_atts.append(
                _file_attachment(p, content_id=pl.Path(p).stem, inline=True)
            )

    # Attach inline images BEFORE setting the body so the cid: refs resolve and
    # aren't stripped by the sanitizer.
    for att in inline_atts:
        graph.request(
            "POST", f"/me/messages/{draft_id}/attachments", account_id, json=att
        )

    # Set the composed body, preserving the source content type (a text-sourced
    # reply can't be promoted to HTML).
    graph.request(
        "PATCH",
        f"/me/messages/{draft_id}",
        account_id,
        json={
            "body": {
                "contentType": "HTML" if source_is_html else "Text",
                "content": full_body,
            }
        },
    )

    # Regular file attachments (large ones go via an upload session).
    if attachments:
        att_paths = [attachments] if isinstance(attachments, str) else attachments
        for file_path in att_paths:
            path = pl.Path(file_path).expanduser().resolve()
            content_bytes = path.read_bytes()
            if len(content_bytes) < 3 * 1024 * 1024:
                graph.request(
                    "POST",
                    f"/me/messages/{draft_id}/attachments",
                    account_id,
                    json=_file_attachment(file_path),
                )
            else:
                graph.upload_large_mail_attachment(
                    draft_id,
                    path.name,
                    content_bytes,
                    account_id,
                    "application/octet-stream",
                )

    result = graph.request("GET", f"/me/messages/{draft_id}", account_id)
    return result if result else {"id": draft_id, "status": "draft"}


@mcp.tool
def list_events(
    account_id: str,
    days_ahead: int = 7,
    days_back: int = 0,
    include_details: bool = True,
) -> list[dict[str, Any]]:
    """List calendar events within specified date range, including recurring event instances"""
    now = dt.datetime.now(dt.timezone.utc)
    start = (now - dt.timedelta(days=days_back)).isoformat()
    end = (now + dt.timedelta(days=days_ahead)).isoformat()

    params = {
        "startDateTime": start,
        "endDateTime": end,
        "$orderby": "start/dateTime",
        "$top": 100,
    }

    if include_details:
        params["$select"] = (
            "id,subject,start,end,location,body,attendees,organizer,isAllDay,recurrence,onlineMeeting,seriesMasterId"
        )
    else:
        params["$select"] = "id,subject,start,end,location,organizer,seriesMasterId"

    # Use calendarView to get recurring event instances
    events = list(
        graph.request_paginated("/me/calendarView", account_id, params=params)
    )

    return events


@mcp.tool
def get_event(event_id: str, account_id: str) -> dict[str, Any]:
    """Get full event details"""
    result = graph.request("GET", f"/me/events/{event_id}", account_id)
    if not result:
        raise ValueError(f"Event with ID {event_id} not found")
    return result


@mcp.tool
def create_event(
    account_id: str,
    subject: str,
    start: str,
    end: str,
    location: str | None = None,
    body: str | None = None,
    attendees: str | list[str] | None = None,
    timezone: str = "UTC",
) -> dict[str, Any]:
    """Create a calendar event"""
    event = {
        "subject": subject,
        "start": {"dateTime": start, "timeZone": timezone},
        "end": {"dateTime": end, "timeZone": timezone},
    }

    if location:
        event["location"] = {"displayName": location}

    if body:
        event["body"] = _email_body(body)

    if attendees:
        attendees_list = [attendees] if isinstance(attendees, str) else attendees
        event["attendees"] = [
            {"emailAddress": {"address": a}, "type": "required"} for a in attendees_list
        ]

    result = graph.request("POST", "/me/events", account_id, json=event)
    if not result:
        raise ValueError("Failed to create event")
    return result


@mcp.tool
def update_event(
    event_id: str, updates: dict[str, Any], account_id: str
) -> dict[str, Any]:
    """Update event properties"""
    formatted_updates = {}

    if "subject" in updates:
        formatted_updates["subject"] = updates["subject"]
    if "start" in updates:
        formatted_updates["start"] = {
            "dateTime": updates["start"],
            "timeZone": updates.get("timezone", "UTC"),
        }
    if "end" in updates:
        formatted_updates["end"] = {
            "dateTime": updates["end"],
            "timeZone": updates.get("timezone", "UTC"),
        }
    if "location" in updates:
        formatted_updates["location"] = {"displayName": updates["location"]}
    if "body" in updates:
        formatted_updates["body"] = _email_body(updates["body"])

    result = graph.request(
        "PATCH", f"/me/events/{event_id}", account_id, json=formatted_updates
    )
    return result or {"status": "updated"}


@mcp.tool
def delete_event(
    account_id: str, event_id: str, send_cancellation: bool = True
) -> dict[str, str]:
    """Delete or cancel a calendar event"""
    if send_cancellation:
        graph.request("POST", f"/me/events/{event_id}/cancel", account_id, json={})
    else:
        graph.request("DELETE", f"/me/events/{event_id}", account_id)
    return {"status": "deleted"}


@mcp.tool
def respond_event(
    account_id: str,
    event_id: str,
    response: str = "accept",
    message: str | None = None,
) -> dict[str, str]:
    """Respond to event invitation (accept, decline, tentativelyAccept)"""
    payload: dict[str, Any] = {"sendResponse": True}
    if message:
        payload["comment"] = message

    graph.request("POST", f"/me/events/{event_id}/{response}", account_id, json=payload)
    return {"status": response}


@mcp.tool
def check_availability(
    account_id: str,
    start: str,
    end: str,
    attendees: str | list[str] | None = None,
) -> dict[str, Any]:
    """Check calendar availability for scheduling"""
    me_info = graph.request("GET", "/me", account_id)
    if not me_info or "mail" not in me_info:
        raise ValueError("Failed to get user email address")
    schedules = [me_info["mail"]]
    if attendees:
        attendees_list = [attendees] if isinstance(attendees, str) else attendees
        schedules.extend(attendees_list)

    payload = {
        "schedules": schedules,
        "startTime": {"dateTime": start, "timeZone": "UTC"},
        "endTime": {"dateTime": end, "timeZone": "UTC"},
        "availabilityViewInterval": 30,
    }

    result = graph.request("POST", "/me/calendar/getSchedule", account_id, json=payload)
    if not result:
        raise ValueError("Failed to check availability")
    return result


@mcp.tool
def list_contacts(account_id: str, limit: int = 50) -> list[dict[str, Any]]:
    """List contacts"""
    params = {"$top": min(limit, 100)}

    contacts = list(
        graph.request_paginated("/me/contacts", account_id, params=params, limit=limit)
    )

    return contacts


@mcp.tool
def get_contact(contact_id: str, account_id: str) -> dict[str, Any]:
    """Get contact details"""
    result = graph.request("GET", f"/me/contacts/{contact_id}", account_id)
    if not result:
        raise ValueError(f"Contact with ID {contact_id} not found")
    return result


@mcp.tool
def create_contact(
    account_id: str,
    given_name: str,
    surname: str | None = None,
    email_addresses: str | list[str] | None = None,
    phone_numbers: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Create a new contact"""
    contact: dict[str, Any] = {"givenName": given_name}

    if surname:
        contact["surname"] = surname

    if email_addresses:
        email_list = (
            [email_addresses] if isinstance(email_addresses, str) else email_addresses
        )
        contact["emailAddresses"] = [
            {"address": email, "name": f"{given_name} {surname or ''}".strip()}
            for email in email_list
        ]

    if phone_numbers:
        if "business" in phone_numbers:
            contact["businessPhones"] = [phone_numbers["business"]]
        if "home" in phone_numbers:
            contact["homePhones"] = [phone_numbers["home"]]
        if "mobile" in phone_numbers:
            contact["mobilePhone"] = phone_numbers["mobile"]

    result = graph.request("POST", "/me/contacts", account_id, json=contact)
    if not result:
        raise ValueError("Failed to create contact")
    return result


@mcp.tool
def update_contact(
    contact_id: str, updates: dict[str, Any], account_id: str
) -> dict[str, Any]:
    """Update contact information"""
    result = graph.request(
        "PATCH", f"/me/contacts/{contact_id}", account_id, json=updates
    )
    return result or {"status": "updated"}


@mcp.tool
def delete_contact(contact_id: str, account_id: str) -> dict[str, str]:
    """Delete a contact"""
    graph.request("DELETE", f"/me/contacts/{contact_id}", account_id)
    return {"status": "deleted"}


@mcp.tool
def list_files(
    account_id: str, path: str = "/", limit: int = 50
) -> list[dict[str, Any]]:
    """List files and folders in OneDrive"""
    endpoint = (
        "/me/drive/root/children"
        if path == "/"
        else f"/me/drive/root:/{path}:/children"
    )
    params = {
        "$top": min(limit, 100),
        "$select": "id,name,size,lastModifiedDateTime,folder,file,@microsoft.graph.downloadUrl",
    }

    items = list(
        graph.request_paginated(endpoint, account_id, params=params, limit=limit)
    )

    return [
        {
            "id": item["id"],
            "name": item["name"],
            "type": "folder" if "folder" in item else "file",
            "size": item.get("size", 0),
            "modified": item.get("lastModifiedDateTime"),
            "download_url": item.get("@microsoft.graph.downloadUrl"),
        }
        for item in items
    ]


@mcp.tool
def get_file(file_id: str, account_id: str, download_path: str) -> dict[str, Any]:
    """Download a file from OneDrive to local path"""
    import subprocess

    metadata = graph.request("GET", f"/me/drive/items/{file_id}", account_id)
    if not metadata:
        raise ValueError(f"File with ID {file_id} not found")

    download_url = metadata.get("@microsoft.graph.downloadUrl")
    if not download_url:
        raise ValueError("No download URL available for this file")

    try:
        subprocess.run(
            ["curl", "-L", "-o", download_path, download_url],
            check=True,
            capture_output=True,
        )

        return {
            "path": download_path,
            "name": metadata.get("name", "unknown"),
            "size_mb": round(metadata.get("size", 0) / (1024 * 1024), 2),
            "mime_type": metadata.get("file", {}).get("mimeType") if metadata else None,
        }
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Failed to download file: {e.stderr.decode()}")


@mcp.tool
def create_file(
    onedrive_path: str, local_file_path: str, account_id: str
) -> dict[str, Any]:
    """Upload a local file to OneDrive"""
    path = pl.Path(local_file_path).expanduser().resolve()
    data = path.read_bytes()
    result = graph.upload_large_file(
        f"/me/drive/root:/{onedrive_path}:", data, account_id
    )
    if not result:
        raise ValueError(f"Failed to create file at path: {onedrive_path}")
    return result


@mcp.tool
def update_file(file_id: str, local_file_path: str, account_id: str) -> dict[str, Any]:
    """Update OneDrive file content from a local file"""
    path = pl.Path(local_file_path).expanduser().resolve()
    data = path.read_bytes()
    result = graph.upload_large_file(f"/me/drive/items/{file_id}", data, account_id)
    if not result:
        raise ValueError(f"Failed to update file with ID: {file_id}")
    return result


@mcp.tool
def delete_file(file_id: str, account_id: str) -> dict[str, str]:
    """Delete a file or folder"""
    graph.request("DELETE", f"/me/drive/items/{file_id}", account_id)
    return {"status": "deleted"}


@mcp.tool
def get_attachment(
    email_id: str, attachment_id: str, save_path: str, account_id: str
) -> dict[str, Any]:
    """Download email attachment to a specified file path"""
    result = graph.request(
        "GET", f"/me/messages/{email_id}/attachments/{attachment_id}", account_id
    )

    if not result:
        raise ValueError("Attachment not found")

    if "contentBytes" not in result:
        raise ValueError("Attachment content not available")

    # Save attachment to file
    path = pl.Path(save_path).expanduser().resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    content_bytes = base64.b64decode(result["contentBytes"])
    path.write_bytes(content_bytes)

    return {
        "name": result.get("name", "unknown"),
        "content_type": result.get("contentType", "application/octet-stream"),
        "size": result.get("size", 0),
        "saved_to": str(path),
    }


@mcp.tool
def search_files(
    query: str,
    account_id: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for files in OneDrive using the modern search API."""
    items = list(graph.search_query(query, ["driveItem"], account_id, limit))

    return [
        {
            "id": item["id"],
            "name": item["name"],
            "type": "folder" if "folder" in item else "file",
            "size": item.get("size", 0),
            "modified": item.get("lastModifiedDateTime"),
            "download_url": item.get("@microsoft.graph.downloadUrl"),
        }
        for item in items
    ]


@mcp.tool
def search_emails(
    query: str,
    account_id: str,
    limit: int = 50,
    folder: str | None = None,
) -> list[dict[str, Any]]:
    """Search emails using the modern search API."""
    if folder:
        # For folder-specific search, use the traditional endpoint
        folder_path = FOLDERS.get(folder.casefold(), folder)
        endpoint = f"/me/mailFolders/{folder_path}/messages"

        params = {
            "$search": f'"{query}"',
            "$top": min(limit, 100),
            "$select": "id,subject,from,toRecipients,receivedDateTime,hasAttachments,body,conversationId,isRead",
        }

        return list(
            graph.request_paginated(endpoint, account_id, params=params, limit=limit)
        )

    return list(graph.search_query(query, ["message"], account_id, limit))


@mcp.tool
def search_events(
    query: str,
    account_id: str,
    days_ahead: int = 365,
    days_back: int = 365,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search calendar events using the modern search API."""
    events = list(graph.search_query(query, ["event"], account_id, limit))

    # Filter by date range if needed
    if days_ahead != 365 or days_back != 365:
        now = dt.datetime.now(dt.timezone.utc)
        start = now - dt.timedelta(days=days_back)
        end = now + dt.timedelta(days=days_ahead)

        filtered_events = []
        for event in events:
            event_start = dt.datetime.fromisoformat(
                event.get("start", {}).get("dateTime", "").replace("Z", "+00:00")
            )
            event_end = dt.datetime.fromisoformat(
                event.get("end", {}).get("dateTime", "").replace("Z", "+00:00")
            )

            if event_start <= end and event_end >= start:
                filtered_events.append(event)

        return filtered_events

    return events


@mcp.tool
def search_contacts(
    query: str,
    account_id: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search contacts. Uses traditional search since unified_search doesn't support contacts."""
    params = {
        "$search": f'"{query}"',
        "$top": min(limit, 100),
    }

    contacts = list(
        graph.request_paginated("/me/contacts", account_id, params=params, limit=limit)
    )

    return contacts


@mcp.tool
def unified_search(
    query: str,
    account_id: str,
    entity_types: list[str] | None = None,
    limit: int = 50,
) -> dict[str, list[dict[str, Any]]]:
    """Search across multiple Microsoft 365 resources using the modern search API

    entity_types can include: 'message', 'event', 'drive', 'driveItem', 'list', 'listItem', 'site'
    If not specified, searches across all available types.
    """
    if not entity_types:
        entity_types = ["message", "event", "driveItem"]

    results = {entity_type: [] for entity_type in entity_types}

    items = list(graph.search_query(query, entity_types, account_id, limit))

    for item in items:
        resource_type = item.get("@odata.type", "").split(".")[-1]

        if resource_type == "message":
            results.setdefault("message", []).append(item)
        elif resource_type == "event":
            results.setdefault("event", []).append(item)
        elif resource_type in ["driveItem", "file", "folder"]:
            results.setdefault("driveItem", []).append(item)
        else:
            results.setdefault("other", []).append(item)

    return {k: v for k, v in results.items() if v}


@mcp.tool
def list_chats(
    account_id: str,
    limit: int = 25,
) -> list[dict[str, Any]]:
    """List the user's Teams chats (1:1 and group), expanded with members.

    Ordered by most-recent activity. Each chat includes id, topic, chatType,
    lastUpdatedDateTime, and a members[] list with displayName + email.
    Use the returned chat id with list_chat_messages.
    """
    params = {
        "$expand": "members",
        "$top": str(min(limit, 50)),
        "$orderby": "lastMessagePreview/createdDateTime desc",
    }
    result = graph.request("GET", "/me/chats", account_id, params=params)
    chats = (result or {}).get("value", []) if result else []
    out = []
    for c in chats:
        out.append(
            {
                "id": c.get("id"),
                "topic": c.get("topic"),
                "chatType": c.get("chatType"),
                "lastUpdatedDateTime": c.get("lastUpdatedDateTime"),
                "webUrl": c.get("webUrl"),
                "members": [
                    {
                        "displayName": m.get("displayName"),
                        "email": m.get("email"),
                        "userId": m.get("userId"),
                    }
                    for m in c.get("members", [])
                ],
            }
        )
    return out


@mcp.tool
def find_chat_with(
    name_or_email: str,
    account_id: str,
    chat_limit: int = 50,
) -> list[dict[str, Any]]:
    """Find Teams chats whose members include someone matching name_or_email
    (case-insensitive substring match on displayName or email).

    Returns the same shape as list_chats, filtered. Useful before list_chat_messages.
    """
    needle = name_or_email.lower()
    # list_chats is wrapped by @mcp.tool (a FunctionTool, not callable) — call the
    # underlying function via .fn. Without this, find_chat_with raises
    # "'FunctionTool' object is not callable".
    all_chats = list_chats.fn(account_id, limit=chat_limit)
    matches = []
    for c in all_chats:
        for m in c.get("members", []):
            dn = (m.get("displayName") or "").lower()
            em = (m.get("email") or "").lower()
            if needle in dn or needle in em:
                matches.append(c)
                break
    return matches


@mcp.tool
def list_chat_messages(
    chat_id: str,
    account_id: str,
    limit: int = 20,
) -> list[dict[str, Any]]:
    """List recent messages from a Teams chat, newest first.

    Returns id, createdDateTime, from (displayName + email), and body.content
    (text/HTML). Use the chat_id from list_chats or find_chat_with.
    """
    params = {"$top": str(min(limit, 50))}
    result = graph.request(
        "GET", f"/chats/{chat_id}/messages", account_id, params=params
    )
    messages = (result or {}).get("value", []) if result else []
    out = []
    for msg in messages:
        sender = (msg.get("from") or {}).get("user") or {}
        out.append(
            {
                "id": msg.get("id"),
                "createdDateTime": msg.get("createdDateTime"),
                "from": {
                    "displayName": sender.get("displayName"),
                    "email": sender.get("email") or sender.get("userPrincipalName"),
                },
                "messageType": msg.get("messageType"),
                "importance": msg.get("importance"),
                "body": (msg.get("body") or {}).get("content"),
                "contentType": (msg.get("body") or {}).get("contentType"),
            }
        )
    return out


@mcp.tool
def send_chat_message(
    chat_id: str,
    message: str,
    account_id: str,
) -> dict[str, Any]:
    """Send a message to a Teams chat (1:1 or group).

    Use the chat_id from list_chats or find_chat_with. Content type is
    auto-detected: HTML passes through; plain text is converted to HTML
    with <br> line breaks (Teams collapses newlines in text-type bodies).
    Uses the Chat.ReadWrite scope already on the token.

    Returns the created message's id, createdDateTime, and webUrl.
    Always confirm the chat's members (via find_chat_with) before sending.
    """
    result = graph.request(
        "POST",
        f"/chats/{chat_id}/messages",
        account_id,
        json={"body": _chat_body(message)},
    )
    result = result or {}
    return {
        "id": result.get("id"),
        "createdDateTime": result.get("createdDateTime"),
        "webUrl": result.get("webUrl"),
        "chatId": result.get("chatId") or chat_id,
    }
