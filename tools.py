import os
import tempfile
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path

from langchain.tools import tool
from ddgs import DDGS
from fpdf import FPDF


@tool
def web_search(query: str) -> str:
    """Search the web for current information. Use this for recent events, news, or facts you're unsure about."""
    try:
        results = DDGS().text(query, max_results=10)
        if not results:
            return "No results found."
        return "\n\n".join(
            f"Title: {r['title']}\nURL: {r['href']}\nSnippet: {r['body']}"
            for r in results
        )
    except Exception as e:
        return f"Search error: {e}"


@tool
def search_images(query: str, max_results: int = 5) -> str:
    """Search the web for images related to a topic.

    Use this tool to find relevant pictures, photos, or diagrams that can be
    embedded into a PDF report. Returns image URLs with metadata.

    Args:
        query: What kind of images to search for (e.g. 'Eiffel Tower at night').
        max_results: Maximum number of image results to return (default 5).
    """
    try:
        results = DDGS().images(query, max_results=max_results)
        if not results:
            return "No image results found."

        output_lines = []
        for i, r in enumerate(results, 1):
            title = r.get("title", "Untitled")
            url = r.get("image", "")
            source = r.get("url", "")
            width = r.get("width", "?")
            height = r.get("height", "?")
            output_lines.append(
                f"{i}. Title: {title}\n"
                f"   Image URL: {url}\n"
                f"   Source: {source}\n"
                f"   Dimensions: {width}x{height}"
            )
        return "\n\n".join(output_lines)
    except Exception as e:
        return f"Image search error: {e}"


class ReportPDF(FPDF):
    """Custom PDF with header and footer for reports."""

    def header(self):
        self.set_font("Helvetica", "B", 12)
        self.cell(0, 10, "Research Report", align="C", new_x="LMARGIN", new_y="NEXT")
        self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
        self.ln(4)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")


def _sanitize(text: str) -> str:
    """Replace characters unsupported by latin-1 (Helvetica) with safe equivalents."""
    result = []
    for ch in text:
        try:
            ch.encode("latin-1")
            result.append(ch)
        except UnicodeEncodeError:
            # Replace common Unicode chars with ASCII equivalents
            replacements = {
                "‘": "'", "’": "'", "“": '"', "”": '"',
                "–": "--", "—": "---", "•": "-",
                "…": "...", " ": " ",
            }
            result.append(replacements.get(ch, "?"))
    return "".join(result)


def _download_image(url: str, dest_dir: str) -> str | None:
    """Download an image from a URL to a local directory. Returns the local path or None."""
    try:
        # Try to get a reasonable file extension from the URL
        path = Path(url.split("?")[0])  # strip query params
        suffix = path.suffix.lower()
        if suffix not in (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"):
            suffix = ".jpg"  # default

        dest = os.path.join(dest_dir, f"img_{hash(url) & 0xFFFF:04x}{suffix}")

        req = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0 (compatible; ScribeMind/1.0)"},
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            with open(dest, "wb") as f:
                f.write(resp.read())
        return dest
    except Exception:
        return None


@tool
def generate_pdf_report(title: str, content: str, images: str = "") -> str:
    """Generate a PDF report and save it to the reports/ directory.

    Use this tool when the user asks you to create a report, write up findings,
    or save search results as a document. Always use after gathering information
    via web_search and optionally search_images for relevant pictures.

    Args:
        title: A concise report title (used as filename and document heading).
        content: The full report body. Use markdown-style formatting:
                 - Use '# Section' for major section headings
                 - Use '## Subsection' for sub-headings
                 - Use '- bullet' for bullet points
                 - Use '[IMG]' as a placeholder where you want an image inserted
                 - Include source URLs inline as plain text
        images: Comma-separated list of image URLs to embed (from search_images).
                Images are placed at [IMG] markers in order, or at the end if no
                markers are present. Example: 'https://example.com/a.jpg, https://example.com/b.png'
    """
    try:
        os.makedirs("reports", exist_ok=True)

        # Parse image URLs
        image_urls = [u.strip() for u in images.split(",") if u.strip()] if images else []

        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        safe_title = "".join(c if c.isalnum() or c in " _-" else "_" for c in title)[:80]
        filename = f"reports/{safe_title}_{timestamp}.pdf"

        # Download images to a temp directory
        tmpdir = tempfile.mkdtemp(prefix="scribemind_img_")
        local_images = []
        for url in image_urls:
            local_path = _download_image(url, tmpdir)
            if local_path:
                local_images.append(local_path)

        pdf = ReportPDF()
        pdf.alias_nb_pages()
        pdf.set_auto_page_break(auto=True, margin=20)
        pdf.add_page()

        # Title
        pdf.set_font("Helvetica", "B", 20)
        pdf.multi_cell(0, 10, _sanitize(title), align="L")
        pdf.ln(2)

        # Metadata line
        pdf.set_font("Helvetica", "I", 9)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 6, f"Generated: {datetime.now().strftime('%B %d, %Y at %H:%M')}", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(6)
        pdf.set_text_color(0, 0, 0)

        # Content — parse sections, bullets, paragraphs, and [IMG] markers
        img_index = 0
        for line in content.split("\n"):
            stripped = line.strip()
            if not stripped:
                pdf.ln(4)
                continue

            if stripped == "[IMG]":
                # Insert the next available image
                if img_index < len(local_images):
                    _place_image(pdf, local_images[img_index], f"Image {img_index + 1}")
                    img_index += 1
                else:
                    pdf.set_font("Helvetica", "I", 9)
                    pdf.set_text_color(150, 150, 150)
                    pdf.cell(0, 6, "[Image placeholder — no image available]", new_x="LMARGIN", new_y="NEXT")
                    pdf.set_text_color(0, 0, 0)
                    pdf.ln(4)
                continue

            if stripped.startswith("# "):
                pdf.set_font("Helvetica", "B", 14)
                pdf.ln(4)
                pdf.multi_cell(0, 8, _sanitize(stripped[2:]), align="L")
                pdf.ln(2)
            elif stripped.startswith("## "):
                pdf.set_font("Helvetica", "B", 12)
                pdf.ln(3)
                pdf.multi_cell(0, 7, _sanitize(stripped[3:]), align="L")
                pdf.ln(1)
            elif stripped.startswith("- ") or stripped.startswith("* "):
                pdf.set_font("Helvetica", "", 10)
                pdf.set_x(pdf.l_margin + 8)
                pdf.multi_cell(pdf.w - pdf.l_margin - pdf.r_margin - 8, 6, "- " + _sanitize(stripped[2:]), align="L")
            else:
                pdf.set_font("Helvetica", "", 10)
                pdf.multi_cell(0, 6, _sanitize(stripped), align="L")

        # Any remaining images not placed via [IMG] markers — append at the end
        if img_index < len(local_images):
            pdf.ln(6)
            pdf.set_font("Helvetica", "B", 12)
            pdf.cell(0, 8, "Related Images", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(6)
            while img_index < len(local_images):
                _place_image(pdf, local_images[img_index], f"Image {img_index + 1}")
                img_index += 1

        pdf.output(filename)

        # Clean up temp images
        for f in os.listdir(tmpdir):
            os.unlink(os.path.join(tmpdir, f))
        os.rmdir(tmpdir)

        return f"PDF report saved to: {os.path.abspath(filename)}"
    except Exception as e:
        return f"PDF generation error: {e}"


def _place_image(pdf: FPDF, image_path: str, caption: str = "") -> None:
    """Place an image into the PDF, sizing it to fit the page width."""
    try:
        # Get image dimensions from the file
        from PIL import Image
        img = Image.open(image_path)
        img_w, img_h = img.size
        img.close()
    except Exception:
        img_w, img_h = 800, 600  # fallback

    # Calculate display size — max width = page width minus margins, max height = 120mm
    max_w = pdf.w - pdf.l_margin - pdf.r_margin
    max_h = 120

    scale = min(max_w / img_w, max_h / img_h, 1.0)
    disp_w = img_w * scale
    disp_h = img_h * scale

    # Center the image
    x = pdf.l_margin + (max_w - disp_w) / 2

    # Check if we need a page break
    if pdf.get_y() + disp_h + 10 > pdf.h - pdf.b_margin:
        pdf.add_page()

    pdf.ln(2)
    pdf.image(image_path, x=x, w=disp_w, h=disp_h)
    pdf.ln(2)

    if caption:
        pdf.set_font("Helvetica", "I", 8)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 5, _sanitize(caption), align="C", new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(0, 0, 0)
    pdf.ln(6)
