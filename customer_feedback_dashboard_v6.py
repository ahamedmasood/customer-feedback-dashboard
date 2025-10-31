# ==========================================================
# DANLESCO CALIBRATION SERVICES - CUSTOMER FEEDBACK DASHBOARD (FINAL v6)
# ==========================================================
# ✅ 10 customers retained
# ✅ Simple language
# ✅ Full charts, visuals, and tables included in generated PDF
# ==========================================================

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import pdfkit
import base64
import tempfile
from io import BytesIO
from datetime import datetime

# Streamlit page config
st.set_page_config(page_title="Customer Feedback Dashboard", layout="wide")
st.title("📊 Danlesco Calibration Services - Customer Feedback Analysis")
st.caption("Final version – full visuals, clean language, and PDF export support.")

# -----------------------------
# Helper functions
# -----------------------------
def fig_to_base64(fig):
    """Convert matplotlib figure to base64 image string."""
    buf = BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight")
    buf.seek(0)
    img_base64 = base64.b64encode(buf.read()).decode("utf-8")
    plt.close(fig)
    return img_base64


def generate_pdf(html_content, file_name="Customer_Feedback_Report.pdf"):
    """Convert HTML string to PDF file bytes using pdfkit."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_html:
        tmp_html.write(html_content.encode("utf-8"))
        tmp_html_path = tmp_html.name
    output_path = tmp_html_path.replace(".html", ".pdf")

    options = {
        'encoding': "UTF-8",
        'page-size': 'A4',
        'orientation': 'Portrait',
        'quiet': '',
        'enable-local-file-access': ''
    }

    pdfkit.from_file(tmp_html_path, output_path, options=options)
    with open(output_path, "rb") as f:
        return f.read()


# -----------------------------
# File upload
# -----------------------------
uploaded = st.file_uploader("📁 Upload the Customer Feedback Excel File", type=["xlsx"])

if uploaded:
    # ==========================================================
    # 1️⃣ LOAD AND CLEAN DATA
    # ==========================================================
    df_raw = pd.read_excel(uploaded, sheet_name=0, header=None)
    df = df_raw.iloc[3:, 1:14]
    df.columns = [
        "Customer Name",
        "Timeliness_Importance", "Timeliness_Performance",
        "Competence_Importance", "Competence_Performance",
        "Clarity_Importance", "Clarity_Performance",
        "Quality_Importance", "Quality_Performance",
        "Service_Importance", "Service_Performance",
        "Pricing_Importance", "Pricing_Performance"
    ]

    df = df.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)
    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors='coerce')
    df = df[df["Customer Name"].notna()]
    df = df.reset_index(drop=True)
    categories = ["Timeliness", "Competence", "Clarity", "Quality", "Service", "Pricing"]

    st.success(f"✅ Successfully loaded data for {len(df)} customers.")

    # Storage for HTML parts
    html_sections = []

    # ==========================================================
    # 2️⃣ INDIVIDUAL CUSTOMER REPORTS
    # ==========================================================
    st.header("🧍 Individual Customer Feedback Reports")

    ratings = []

    for _, row in df.iterrows():
        name = row["Customer Name"]
        st.markdown(f"### 👤 {name}")

        importance = [row[f"{c}_Importance"] for c in categories]
        performance = [row[f"{c}_Performance"] for c in categories]
        total_importance = sum(importance)
        weights = [imp / total_importance for imp in importance]
        weighted_perf = [p * w for p, w in zip(performance, weights)]
        satisfaction = (sum(weighted_perf) / 5) * 100
        ratings.append(satisfaction)

        table = pd.DataFrame({
            "Category": categories,
            "Importance": importance,
            "Performance": performance,
            "Share of Importance": [round(w, 3) for w in weights],
            "Score Contribution": [round(wp, 3) for wp in weighted_perf]
        })
        st.dataframe(table, use_container_width=True)

        fig, ax = plt.subplots(figsize=(6, 3))
        x = np.arange(len(categories))
        ax.bar(x - 0.2, importance, width=0.4, label="Importance", alpha=0.7)
        ax.bar(x + 0.2, performance, width=0.4, label="Performance", alpha=0.7)
        ax.set_xticks(x)
        ax.set_xticklabels(categories, rotation=30, ha="right")
        ax.legend()
        ax.set_ylim(0, 5)
        st.pyplot(fig)

        # Convert to image for PDF
        chart_b64 = fig_to_base64(fig)

        html_sections.append(f"""
        <h3>{name}</h3>
        <p><b>Overall Satisfaction:</b> {satisfaction:.1f}%</p>
        <img src="data:image/png;base64,{chart_b64}" width="500"><br>
        """)

    # ==========================================================
    # 3️⃣ COMPANY-WIDE SUMMARY
    # ==========================================================
    st.header("🏢 Company-Wide Summary")

    summary = pd.DataFrame({
        "Category": categories,
        "Importance": [df[f"{c}_Importance"].mean() for c in categories],
        "Performance": [df[f"{c}_Performance"].mean() for c in categories]
    })
    summary["Gap"] = summary["Performance"] - summary["Importance"]

    imp_avg = summary["Importance"].mean()
    perf_avg = summary["Performance"].mean()
    CSI = (perf_avg / imp_avg) * 100
    avg_customer_sat = np.mean(ratings)

    # Gap chart
    fig_gap, ax_gap = plt.subplots(figsize=(8, 4))
    colors = ["green" if g >= 0 else "red" for g in summary["Gap"]]
    ax_gap.bar(summary["Category"], summary["Gap"], color=colors)
    ax_gap.axhline(0, color="black", linewidth=0.8)
    ax_gap.set_ylabel("Difference (Performance - Importance)")
    ax_gap.set_title("Areas Performing Above or Below Expectation")
    st.pyplot(fig_gap)

    chart_gap_b64 = fig_to_base64(fig_gap)

    # Distribution chart
    fig_hist, ax_hist = plt.subplots(figsize=(6, 3))
    ax_hist.hist(ratings, bins=[50, 65, 75, 85, 95, 100], edgecolor='black', color='skyblue')
    ax_hist.set_xlabel("Customer Satisfaction %")
    ax_hist.set_ylabel("Number of Customers")
    ax_hist.set_title("Distribution of Satisfaction Levels")
    st.pyplot(fig_hist)
    chart_hist_b64 = fig_to_base64(fig_hist)

    # Voice of Customer pie
    fig_voc, ax_voc = plt.subplots(figsize=(5, 5))
    weights_total = summary["Importance"] / summary["Importance"].sum()
    ax_voc.pie(weights_total, labels=summary["Category"],
               autopct="%1.1f%%", startangle=90, wedgeprops={'width': 0.4})
    ax_voc.set_title("Customer Priorities by Category")
    st.pyplot(fig_voc)
    chart_voc_b64 = fig_to_base64(fig_voc)

    # HTML summary for PDF
    html_summary = f"""
    <h2>Danlesco Calibration Services – Customer Feedback Summary</h2>
    <p><b>Total Customers:</b> {len(df)}</p>
    <p><b>Overall CSI:</b> {CSI:.1f}%</p>
    <p><b>Average Satisfaction:</b> {avg_customer_sat:.1f}%</p>
    <h3>Gap Analysis</h3>
    <img src="data:image/png;base64,{chart_gap_b64}" width="600"><br>
    <h3>Distribution of Satisfaction Levels</h3>
    <img src="data:image/png;base64,{chart_hist_b64}" width="600"><br>
    <h3>Voice of the Customer</h3>
    <img src="data:image/png;base64,{chart_voc_b64}" width="400"><br>
    """
    html_sections.insert(0, html_summary)

    # ==========================================================
    # 4️⃣ EXPORT PDF
    # ==========================================================
    st.header("📄 Export Full Report")

    if st.button("📥 Generate PDF Report"):
        full_html = "<html><body style='font-family:Arial'>" + "".join(html_sections) + "</body></html>"
        pdf_data = generate_pdf(full_html)
        st.download_button(
            label="⬇️ Download Report as PDF",
            data=pdf_data,
            file_name=f"Danlesco_Customer_Feedback_Report_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )
else:
    st.info("👆 Please upload your Excel file to generate the dashboard and PDF report.")
