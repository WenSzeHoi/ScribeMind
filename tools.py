import os
from datetime import datetime

from langchain.tools import tool
from ddgs import DDGS
from fpdf import FPDF


@tool
def web_search(query: str) -> str:
    """Search the web for current information. Use this for recent events, news, or facts you're unsure about."""
    try:
        results = DDGS().text(query, max_results=9)
        if not results:
            return "No results found."
        return "\n\n".join(
            f"Title: {r['title']}\nURL: {r['href']}\nSnippet: {r['body']}"
            for r in results
        )
    except Exception as e:
        return f"Search error: {e}"


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


@tool
def generate_pdf_report(title: str, content: str) -> str:
    """Generate a PDF report and save it to the reports/ directory.

    Use this tool when the user asks you to create a report, write up findings,
    or save search results as a document. Always use after gathering information via web_search.

    Args:
        title: A concise report title (used as filename and document heading).
        content: The full report body. Use markdown-style formatting:
                 - Use '# Section' for major section headings
                 - Use '## Subsection' for sub-headings
                 - Use '- bullet' for bullet points
                 - Include source URLs inline as plain text
    """
    try:
        os.makedirs("reports", exist_ok=True)

        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        safe_title = "".join(c if c.isalnum() or c in " _-" else "_" for c in title)[:80]
        filename = f"reports/{safe_title}_{timestamp}.pdf"

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

        # Content — parse sections, bullets, and paragraphs
        for line in content.split("\n"):
            stripped = line.strip()
            if not stripped:
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

        pdf.output(filename)
        return f"PDF report saved to: {os.path.abspath(filename)}"
    except Exception as e:
        return f"PDF generation error: {e}"
