"""
Generate the two architecture diagrams for the DSM-NIS Booking ARB, styled to
match the reference Tech Assurance Group template (Akasha) look & feel.

  image1  -> Logical architecture (source triggers -> engine -> outcome)
  image2  -> Layered technical architecture (presentation/API/logic/data + cross-cutting)
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrow
import os

OUT = os.path.dirname(os.path.abspath(__file__))

# ---- palette (mirrors the reference template) ----
DARK_TEAL = "#0f5b52"
GREY = "#e6e6e6"
CYAN_BG = "#cfeef7"
CYAN_BOX = "#eaf7fb"
CYAN_HDR = "#1f9bc0"
GREEN_BG = "#dcedc8"
CARD_GREY = "#d9d9d9"
RED_BG = "#f8d0cf"
YELLOW_BG = "#ffe0a3"
BLUE_TXT = "#1f6fc0"


def rbox(ax, x, y, w, h, text, fc, ec="none", tc="black", fs=9, bold=False, rounding=0.02):
    box = FancyBboxPatch((x, y), w, h,
                         boxstyle=f"round,pad=0.005,rounding_size={rounding}",
                         linewidth=1, edgecolor=ec, facecolor=fc)
    ax.add_patch(box)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center",
            fontsize=fs, color=tc, weight="bold" if bold else "normal", wrap=True)


def arrow(ax, x1, y1, x2, y2):
    ax.add_patch(FancyArrow(x1, y1, x2 - x1, y2 - y1, width=0.002,
                            head_width=0.012, head_length=0.02,
                            length_includes_head=True, color="#444444"))


# =====================================================================
# IMAGE 1 - Logical architecture flow
# =====================================================================
def make_logical():
    fig, ax = plt.subplots(figsize=(11, 5.6))
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    # Top chevron sources
    sources = [
        ("Outlook Mailbox", "DSM-NIS-BOOKING emails,\napprovals, D/C replies"),
        ("HTML Table Parser", "Metadata + SPV\ncharge extraction"),
        ("SharePoint / Excel", "Vendor master, output\ntracker, booking folders"),
        ("NIS Fiori Portal", "Non-PO vendor payment\nbooking (+ve amounts)"),
        ("SAP GUI (ZDCC)", "Debit/Credit checklist\n(-ve amounts)"),
    ]
    n = len(sources)
    gap = 0.012
    w = (1 - (n - 1) * gap - 0.02) / n
    x0 = 0.01
    for i, (title, sub) in enumerate(sources):
        x = x0 + i * (w + gap)
        # chevron header
        rbox(ax, x, 0.88, w, 0.08, title, DARK_TEAL, tc="white", fs=9, bold=True)
        # detail card
        rbox(ax, x, 0.70, w, 0.16, sub, GREY, fs=8)
        arrow(ax, x + w / 2, 0.70, x + w / 2, 0.615)

    # Central engine band
    rbox(ax, 0.06, 0.30, 0.88, 0.30, "", CYAN_BG, rounding=0.03)
    rbox(ax, 0.34, 0.545, 0.32, 0.06, "DSM-NIS Booking Automation Engine",
         CYAN_HDR, tc="white", fs=11, bold=True)

    engines = [
        ("Email Classification &\nRouting", "Subject triggers, +ve/-ve/0\namount classification per SPV row"),
        ("Booking Orchestration", "Parallel NIS booking, cost-center\ngrouping, SAP checklist creation"),
        ("AI Invoice Extraction &\nApproval Workflow", "Azure OpenAI invoice read,\nHOD/FC approval, PDF merge"),
    ]
    ew = 0.27
    ex0 = 0.075
    egap = (0.88 - 0.03 - 3 * ew) / 2
    for i, (t, s) in enumerate(engines):
        x = ex0 + i * (ew + egap)
        rbox(ax, x, 0.46, ew, 0.055, t, CYAN_HDR, tc="white", fs=8.5, bold=True)
        rbox(ax, x, 0.335, ew, 0.12, s, CYAN_BOX, ec="#8fd3e6", fs=7.5)

    arrow(ax, 0.5, 0.30, 0.5, 0.235)

    # Persistence / tracking band
    rbox(ax, 0.15, 0.16, 0.70, 0.075,
         "PostgreSQL State Store   (email_runs | booking_rows | run_steps | debit_credit_*)",
         DARK_TEAL, tc="white", fs=8.5, bold=True)
    arrow(ax, 0.5, 0.16, 0.5, 0.10)

    # Outcome band
    rbox(ax, 0.12, 0.015, 0.76, 0.075,
         "Fully booked DSM-NIS transactions with merged evidence PDFs, updated tracker, "
         "and stage-wise completion notifications",
         GREEN_BG, ec="#a5c884", fs=8.5)

    plt.savefig(os.path.join(OUT, "diagram_logical.png"), dpi=200,
                bbox_inches="tight", facecolor="white")
    plt.close()


# =====================================================================
# IMAGE 2 - Layered technical architecture
# =====================================================================
def make_layered():
    fig, ax = plt.subplots(figsize=(11, 5.8))
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

    layer_x = 0.02
    layer_w = 0.68

    def layer(y, h, title, boxes, box_fc, box_ec, dash, box_h=None, box_top=None):
        # dashed container
        cont = FancyBboxPatch((layer_x, y), layer_w, h,
                              boxstyle="round,pad=0.005,rounding_size=0.02",
                              linewidth=1.3, edgecolor=box_ec, facecolor="none",
                              linestyle=dash)
        ax.add_patch(cont)
        ax.text(layer_x + 0.01, y + h - 0.028, title, ha="left", va="center",
                fontsize=9.5, weight="bold", color="#333333")
        bn = len(boxes)
        if bn == 0:
            return
        bgap = 0.015
        bw = (layer_w - 0.03 - (bn - 1) * bgap) / bn
        bx0 = layer_x + 0.015
        bh = box_h if box_h is not None else (h - 0.055)
        by = box_top if box_top is not None else (y + 0.012)
        for i, b in enumerate(boxes):
            bx = bx0 + i * (bw + bgap)
            rbox(ax, bx, by, bw, bh, b, box_fc, ec=box_ec, fs=7.8)

    # Layer 1 - Presentation / Interface
    layer(0.80, 0.16, "1. Interface Layer  (how the system is triggered & reports)",
          ["Outlook Email\n(trigger + notifications)", "Windows Task Scheduler\n(3x daily runs)",
           "SharePoint Tracker\n& Booking Folders"],
          CARD_GREY, "#9e9e9e", "-")

    # Layer 2 - Integration / Identity
    layer(0.60, 0.16, "2. Integration & Identity Layer  (connectors, auth)",
          ["Microsoft Graph API\n(MSAL / OAuth2)", "Azure AD App\nRegistrations", "Azure OpenAI\nREST API"],
          YELLOW_BG, "#e0a83a", "--")

    # Layer 3 - Business logic / Application
    layer(0.34, 0.22, "3. Business Logic / Application Layer  (automation engines)",
          ["Email Classifier\n& Router", "NIS Booking Engine\n(Playwright / Edge)",
           "SAP Automation Engine\n(win32com COM)"],
          RED_BG, "#d9534f", ":", box_h=0.07, box_top=0.455)
    # second row of engines inside layer 3
    for i, b in enumerate(["Invoice Extractor\n(pdfplumber + GPT)",
                            "Approval Workflow\n(HOD / FC)",
                            "PDF Merge &\nDebit/Credit Flow"]):
        bw = (layer_w - 0.03 - 2 * 0.015) / 3
        bx = layer_x + 0.015 + i * (bw + 0.015)
        rbox(ax, bx, 0.355, bw, 0.07, b, RED_BG, ec="#d9534f", fs=7.6)

    # Layer 4 - Data
    layer(0.14, 0.16, "4. Data Layer  (state, persistence, logging)",
          ["Azure PostgreSQL\n(pipeline state)", "SharePoint / Excel\n(vendor + tracker)",
           "Per-SPV File Logs\n(log_manager)"],
          GREEN_BG, "#7cb342", "-")

    # Cross-cutting panel
    cc_x = 0.74
    rbox(ax, cc_x, 0.14, 0.24, 0.82, "", "#efefef", ec="#cccccc", rounding=0.02)
    ax.text(cc_x + 0.12, 0.90, "Cross-Cutting\nServices", ha="center", va="center",
            fontsize=11, weight="bold", color="#333333")
    cc = [
        ("Identity & Security", "Azure AD, OAuth2 client\ncredentials, RSA/TOTP OTP,\nsecrets via env / .env"),
        ("Resilience & Resume", "Step-level DB tracking,\ncrash-resume, idempotency,\nretry limits"),
        ("Observability", "Per-SPV logs, failure email\nwith LLM explanation"),
        ("DevOps", "GitLab source control,\nWindows VM deployment"),
    ]
    yy = 0.80
    for title, body in cc:
        ax.text(cc_x + 0.12, yy, title, ha="center", va="center",
                fontsize=9, weight="bold", color=BLUE_TXT)
        ax.text(cc_x + 0.12, yy - 0.045, body, ha="center", va="center",
                fontsize=7.2, color="#555555")
        yy -= 0.155

    # connector arrows between layers
    for y in (0.80, 0.60, 0.34):
        arrow(ax, layer_x + layer_w / 2, y, layer_x + layer_w / 2, y - 0.008)

    plt.savefig(os.path.join(OUT, "diagram_layered.png"), dpi=200,
                bbox_inches="tight", facecolor="white")
    plt.close()


if __name__ == "__main__":
    make_logical()
    make_layered()
    print("diagrams written to", OUT)
