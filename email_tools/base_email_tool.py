import io
from email.mime.text import MIMEText  # fixed module path
from email.utils import formatdate
from email.header import Header
import pystache
import html
import logging

from upload_tools import upload_file
from template_utils import find_email_template

logger = logging.getLogger(__name__)


def _load_template() -> str:
    """Load the email HTML template from custom/default template directories.

    Priority:
      1. custom_email_template.html (or /app/custom_templates in production)
      2. default_email_template.html (or /app/default_templates in production)

    Raises FileNotFoundError if none exist.
    """
    path = find_email_template()
    if not path:
        logger.error("Email template not found: tried custom_email_template.html and default_email_template.html")
        raise FileNotFoundError(
            "Email template not found: tried custom_email_template.html and default_email_template.html"
        )
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        logger.error("Failed to read email template at %s: %s", path, e, exc_info=True)
        raise


def create_eml(to=None, cc=None, bcc=None, re=None, content=None, priority="normal", language="cs-CZ", file_name=None):
    """Create an unsent email draft (EML) using a Mustache HTML template.

    Template variables:
      {{language}}  - inserted into lang attributes (sanitized)
      {{subject}}   - inserted (HTML-escaped) into <title>
      {{{content}}} - raw HTML fragment for email body (caller restricted to allowed tags)
    """

    # Validate priority
    if priority and priority.lower() not in ["low", "normal", "high"]:
        raise ValueError("Priority must be 'low', 'normal', or 'high'")

    if not content:
        raise ValueError("Email content is required")
    if not re:
        raise ValueError("Email subject is required")

    buffer = None
    try:
        template_html = _load_template()

        # Prepare context
        safe_language = (language or "").replace('"', '').replace("'", '')
        escaped_subject = html.escape(re or "")

        renderer = pystache.Renderer(escape=lambda u: u)  # We'll manually escape where needed
        context = {
            "language": safe_language,   # safe for attribute insertion
            "subject": escaped_subject,  # already escaped
            "content": content,          # inserted unescaped via triple braces {{{content}}}
        }
        complete_html = renderer.render(template_html, context)

        msg = MIMEText(complete_html, 'html', 'utf-8')
        # Ensure proper encoding (base64 avoids quoted-printable soft breaks generating '=')
        if 'Content-Transfer-Encoding' in msg:
            msg.replace_header('Content-Transfer-Encoding', 'base64')
        else:
            msg.add_header('Content-Transfer-Encoding', 'base64')

        if to:
            msg['To'] = ", ".join(to)
        if cc:
            msg['Cc'] = ", ".join(cc)
        if bcc:
            msg['Bcc'] = ", ".join(bcc)

        msg['Subject'] = Header(re, 'utf-8')
        msg['Date'] = formatdate(localtime=True)
        msg['Content-Language'] = safe_language
        msg['Accept-Language'] = safe_language

        if priority.lower() == 'high':
            msg['X-Priority'] = '1 (Highest)'
            msg['X-MSMail-Priority'] = 'High'
            msg['Importance'] = 'High'
        elif priority.lower() == 'low':
            msg['X-Priority'] = '5 (Lowest)'
            msg['X-MSMail-Priority'] = 'Low'
            msg['Importance'] = 'Low'

        msg['X-Unsent'] = '1'

        buffer = io.BytesIO()
        msg_bytes = msg.as_bytes()
        buffer.write(msg_bytes)
        buffer.seek(0)

        return upload_file(buffer, "eml", filename=file_name)
    except Exception as e:
        logger.error("Failed to create email draft: %s", e, exc_info=True)
        raise RuntimeError(f"Failed to create email draft: {e}") from e
    finally:
        if buffer:
            buffer.close()
