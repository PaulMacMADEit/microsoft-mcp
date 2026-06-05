"""Unit tests for reply-body composition (no network / no auth)."""
from microsoft_mcp.tools import _compose_reply_body, _file_attachment


def test_html_inserts_reply_inside_body_before_quote():
    quoted = "<html><head></head><body><div>QUOTED HISTORY</div></body></html>"
    out = _compose_reply_body("<p>MY REPLY</p>", quoted, source_is_html=True)
    # reply text comes before the quoted history
    assert out.index("MY REPLY") < out.index("QUOTED HISTORY")
    # and it was inserted *inside* the existing <body>, not before <html>
    assert out.startswith("<html><head></head><body><p>MY REPLY</p>")
    assert out.count("<body") == 1


def test_html_with_body_attributes():
    quoted = '<html><body dir="ltr" style="x"><p>Q</p></body></html>'
    out = _compose_reply_body("<p>R</p>", quoted, source_is_html=True)
    assert '<body dir="ltr" style="x"><p>R</p>' in out
    assert out.index("R") < out.index("Q")


def test_html_without_body_tag_prepends():
    out = _compose_reply_body("<p>R</p>", "<div>Q</div>", source_is_html=True)
    assert out.index("R") < out.index("Q")
    assert "<br><br>" in out


def test_plaintext_joins_with_blank_line():
    out = _compose_reply_body("hello", "QUOTED", source_is_html=False)
    assert out == "hello\n\nQUOTED"


def test_file_attachment_inline_sets_cid_and_flag(tmp_path):
    p = tmp_path / "panel.png"
    p.write_bytes(b"\x89PNG\r\n\x1a\n fake")
    att = _file_attachment(str(p), content_id="panel", inline=True)
    assert att["@odata.type"] == "#microsoft.graph.fileAttachment"
    assert att["name"] == "panel.png"
    assert att["contentId"] == "panel"
    assert att["isInline"] is True
    assert "contentBytes" in att and isinstance(att["contentBytes"], str)


def test_file_attachment_plain_has_no_cid():
    # use this very test file as a readable path
    att = _file_attachment(__file__)
    assert "contentId" not in att
    assert "isInline" not in att
