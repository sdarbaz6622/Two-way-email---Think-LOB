"""Lenovo CEC - 2-Way Email Ageing Report Generator.

Flow: upload the raw Excel -> the app maps agents to Team Leads, buckets the
ageing (01 Day, 02 Days, 03 Days, 04-05 Days, 06-10 Days, 11-20 Days,
21-30 Days, 31-60 Days, 60+ Days), builds a formatted Excel workbook and an
Outlook-ready email, and saves that email (To/CC filled, Excel attached) as a
draft in your Exchange mailbox. The draft appears in the Drafts folder of the
new Outlook, Outlook on the web and classic Outlook. Nothing is sent.

Run with:  streamlit run app.py
"""

import os
import re
import subprocess
import tempfile
from email.message import EmailMessage
from html import escape
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ============================================================
# HEADCOUNT / TEAM LEAD MAPPING
# ============================================================

hc_dict = {
    "Alex Xavier": "Alex Xavier",
    "Sumith Kumar Dey": "Sreejith Menon (Support Function)",
    "SHIVAKUMAR SHIRSHI": "A Afrid",
    "Amog K": "Sreejith Menon (Support Function)",
    "M Mohammed Khaja": "Sreejith Menon (Support Function)",
    "Illahin Naaz": "A Afrid",
    "Nikitha Kumari": "A Afrid",
    "Jisharaj K": "Sreejith Menon (Support Function)",
    "Aamir Hassan": "A Afrid",
    "Varshikey Tiwari": "A Afrid",
    "Mohan Kumar": "A Afrid",
    "Rohith N Shetty": "A Afrid",
    "Madhusudan Singh": "A Afrid",

    "Priyanshu kumar Singh": "Sreejith Menon",

    "Aman Kumar": "Augustine Joseph",
    "Ashish Kashyap": "Augustine Joseph",
    "Augustine Joseph": "Augustine Joseph",
    "Bijit Sen": "Augustine Joseph",
    "Harshith Singh N R": "Augustine Joseph",
    "Prakash Ambli": "Augustine Joseph",
    "Shaik Imran": "Augustine Joseph",
    "Shaik Mubarak": "Augustine Joseph",
    "Snehan Thejaswi T": "Augustine Joseph",

    "Devapatla Swetha": "Avrel Fernandez",
    "Ashish Kumar": "Avrel Fernandez",
    "Avrel Fernandez": "Avrel Fernandez",
    "Chandan M B": "Avrel Fernandez",
    "Chandana S": "Avrel Fernandez",
    "Mdshamshad Iqbal": "Avrel Fernandez",
    "Nivedita Kamate": "Avrel Fernandez",
    "Suresh Panda": "Avrel Fernandez",

    "Dinesh S": "Deepak Ghosh (Workstation)",
    "Chandra Prakash Singh": "Deepak Ghosh (Workstation)",
    "Deepak Ghosh": "Deepak Ghosh (Workstation)",
    "Hitesh Diliprao Deshmukh": "Deepak Ghosh (Workstation)",
    "Syed Hussain": "Deepak Ghosh (Workstation)",
    "Subham Prasad": "Deepak Ghosh (Workstation)",
    "Subrata Jana": "Deepak Ghosh (Workstation)",
    "Vinay Kumar Sharma": "Deepak Ghosh (Workstation)",

    "Anand Divatai": "Ram Mohan Singh",
    "Bijay Chettri": "Ram Mohan Singh",
    "Husna C": "Ram Mohan Singh",
    "Jyoti Nandini": "Ram Mohan Singh",
    "Manjushree Manjushree": "Ram Mohan Singh",
    "Miyani Smit": "Ram Mohan Singh",
    "Ram Mohan Singh": "Ram Mohan Singh",
    "Roshan Kumar": "Ram Mohan Singh",
    "Sha Ghousia": "Ram Mohan Singh",

    "Ashish Seet": "Seema Lal",
    "Debasish Panda": "Seema Lal",
    "Debasmita Samal": "Seema Lal",
    "Geetika Garnaik": "Seema Lal",
    "Paramita Ghosh": "Seema Lal",
    "Romeo Fernandez": "Seema Lal",
    "Rushikesh Powar": "Seema Lal",
    "Sandeep Das": "Seema Lal",
    "Sarvesh Singh": "Seema Lal",
    "Seema Lal": "Seema Lal",
    "Shafiulla Sameer": "Seema Lal",
    "Velisetty Kumar": "Seema Lal",

    "Anees Basha": "Manikanta KR",
    "Arshit Choubey": "Manikanta KR",
    "Imad Ulla Khan": "Manikanta KR",
    "Kanire Ravindra": "Manikanta KR",
    "Muhammed Ameen Sheikh": "Manikanta KR",
    "Naveen R": "Manikanta KR",
    "Praveen Kumar": "Manikanta KR",
    "Puranik Tippanna": "Manikanta KR",
    "Shrikant S": "Manikanta KR",
    "Siva Venkatraman": "Manikanta KR",
    "Solomon D": "Manikanta KR",

    "Aliya H": "Sreejith Menon",
    "Preeti Barik": "Sreejith Menon",
    "Priyambada Mohapatra": "Sreejith Menon",
    "Ram Govind": "Sreejith Menon",
    "Richa Panwar": "Sreejith Menon",
    "Rohini Bahubali Munnoli": "Sreejith Menon",
    "Sreejith Menon": "Sreejith Menon",
    "Sudharani Sudharani": "Sreejith Menon",

    "CHAITY DEBNATH": "Syed Arbaz",
    "Mitali Basumatary": "Syed Arbaz",
    "Mohammed Abdul Razzaq": "Syed Arbaz",
    "Shaik Ruhina Arshiya": "Syed Arbaz",
    "Shifa Mohammedi": "Syed Arbaz",
    "Syed Arbaz": "Syed Arbaz",
    "Vaishnavi Vaman Kulkarni": "Syed Arbaz",

    "A r riwan Ahmed": "Vikas B Bhovi",
    "ABHISHEK KUMAR": "Vikas B Bhovi",
    "Govindsa Ankita": "Vikas B Bhovi",
    "Mohammed Saqib": "Vikas B Bhovi",
    "Saniya Saniya": "Vikas B Bhovi",
    "Sneha Sneha": "Vikas B Bhovi",
    "Sneha Patil": "Vikas B Bhovi",
    "Vikas B Bhovi": "Vikas B Bhovi",
    "Vikrant Kumar": "Vikas B Bhovi",
    "Vinay Kumar Joshi": "Vikas B Bhovi",

    "Akshay Anil Mardhekar": "Think (L1.5)",
    "Amol Shioshankar Bhongade": "Think (L1.5)",
    "Arul S": "Think (L1.5)",
    "Hemavathi S": "Think (L1.5)",
    "Jagatheesh A": "Think (L1.5)",
    "Kiran J S": "Think (L1.5)",
    "Liyakathali Mohameth Khan": "Think (L1.5)",
    "Marwan Fahad Faiz Jabbar": "Think (L1.5)",
    "Md Khaja": "Think (L1.5)",
    "Mohammad Sajid": "Think (L1.5)",
    "Narsim Raj T S": "Think (L1.5)",
    "Naveen shivram Naik": "Think (L1.5)",
    "Raju Mondal": "Think (L1.5)",
    "Rammilan Pandit": "Think (L1.5)",
    "Santosh Kumar Biredar": "Think (L1.5)",
    "Sathish Kumar": "Think (L1.5)",
    "Subhani M": "Think (L1.5)",
    "Suresh R": "Think (L1.5)",
    "Vicky D": "Think (L1.5)",
    "Vimal Kumar V": "Think (L1.5)",

    "A Afrid": "A Afrid",

    "Riyaz Pasha": "Deepak Ghosh (Billable)",
    "Rahul Bhanumurthy": "Deepak Ghosh (Billable)",
    "Ramkrishnadas Mondal": "Deepak Ghosh (Billable)",
    "Ayesha Sultana": "Deepak Ghosh (Billable)",

    "Subrat Muduli": "Sreejith Menon (Support Function)",
    "Gulam Hasan": "Sreejith Menon (Support Function)",

    "Chahat Ramani": "A Afrid",
    "Faiz Ahmed": "A Afrid",
    "Steven Mario Pinto": "A Afrid",
    "Shaik Nasreen": "A Afrid",
    "Deepa S Dyamangoudar": "A Afrid",
    "Thowqeer Ahamed": "A Afrid",

    "Kanishka Kanishka": "Augustine Joseph",
    "Mohammed Roushan": "Augustine Joseph",
    "Sushmashri Sushmashri": "Augustine Joseph",
    "Anil Patode": "Augustine Joseph",

    "Suryakant Sahoo": "Avrel Fernandez",
    "Jashwanth Raj": "Avrel Fernandez",
    "Ratnashree P Mallshetty": "Avrel Fernandez",
    "Singitham Hanish": "Avrel Fernandez",
    "Dibyajyoti Panigrahi": "Avrel Fernandez",
    "Nivedita S Ganachari": "Avrel Fernandez",

    "Nikhil Kumar": "Ram Mohan Singh",
    "Mohammed Hussain": "Ram Mohan Singh",
    "Aradhya Gupta": "Ram Mohan Singh",
    "Suhas AB": "Ram Mohan Singh",

    "Vishal Pandey": "Seema Lal",
    "Pragati Priya": "Seema Lal",
    "Dhruv Mishra": "Seema Lal",

    "Veeresh Biradar": "Sreejith Menon",
    "Mohd Kaif": "Sreejith Menon",
    "Arvind Swami": "Sreejith Menon",

    "Annu Priya": "Syed Arbaz",
    "Gokul Anand": "Syed Arbaz",

    "Rajneesh": "Vikas B Bhovi",
    "Priyanka Hanje": "Vikas B Bhovi",
    "Swayam Panda": "Vikas B Bhovi",
    "Sakshi Bharti": "Vikas B Bhovi",

    "Subhranshu Mishra": "Sreejith Menon",
    "Pratik Kittur": "Sreejith Menon",
}

hc_lower = {str(k).strip().lower(): v for k, v in hc_dict.items()}

# One display name per agent, so "ILLAHIN NAAZ" and "Illahin Naaz" are the
# same person and show up in the same tidy case.
canonical_name = {str(k).strip().lower(): str(k).strip() for k in hc_dict}


# ============================================================
# EMAIL RECIPIENTS (defaults, editable in the app)
# ============================================================

DEFAULT_TO = """
ssingh42@lenovo.com; rpowar@lenovo.com; pghosh2@lenovo.com; ggarnaik@lenovo.com;
dsamal1@lenovo.com; smishra22@lenovo.com; psingh44@lenovo.com; asultana@lenovo.com;
sghousia@lenovo.com; aahmed9@lenovo.com; akumar127@lenovo.com; msaqib@lenovo.com;
rkumar79@lenovo.com; ssaniya@lenovo.com; msmit@lenovo.com; ssneha1@lenovo.com;
mmanjushree@lenovo.com; hc4@lenovo.com; bchettri@lenovo.com; cdebnath@lenovo.com;
mbasumatary@lenovo.com; mrazzaq@lenovo.com; pkulkarni5@lenovo.com; rbhanumurthy@lenovo.com;
rmondal4@lenovo.com; sarshiya@lenovo.com; smohammedi@lenovo.com; vkulkarni4@lenovo.com;
rpanwar@lenovo.com; pmohapatra@lenovo.com; pkittur@lenovo.com; ssameer@lenovo.com;
sdas7@lenovo.com; rfernandez4@lenovo.com; dpanda@lenovo.com; aseet@lenovo.com;
vtiwari5@lenovo.com; sshirshi@lenovo.com; rshetty5@lenovo.com; nkumari7@lenovo.com;
msingh28@lenovo.com; inaaz@lenovo.com; ahassan7@lenovo.com; st13@lenovo.com;
smubarak@lenovo.com; simran5@lenovo.com; pambli@lenovo.com; hr6@lenovo.com;
cikhar@lenovo.com; bsen@lenovo.com; avamshi@lenovo.com; akumar72@lenovo.com;
spanda2@lenovo.com; nkamate@lenovo.com; dswetha@lenovo.com; miqbal@lenovo.com;
cs3@lenovo.com; cb2@lenovo.com; akumar70@lenovo.com; svenkatraman@lenovo.com;
ss38@lenovo.com; ptippanna@lenovo.com; pkumar33@lenovo.com; nr3@lenovo.com;
msheikh@lenovo.com; kravindra1@lenovo.com; ikhan10@lenovo.com; achoubey2@lenovo.com;
abasha1@lenovo.com; vsharma11@lenovo.com; sjana1@lenovo.com; sprasad8@lenovo.com;
shussain4@lenovo.com; hdeshmukh@lenovo.com; csingh3@lenovo.com; ds6@lenovo.com;
ja1@lenovo.com; lkhan@lenovo.com; sm21@lenovo.com; mkhaja1@lenovo.com;
ghasan@lenovo.com; smuduli1@lenovo.com; snasreen@lenovo.com; mingin@lenovo.com;
apatode@lenovo.com; gmahesh@lenovo.com; skeshwapur@lenovo.com; sbharti1@lenovo.com;
dpanigrahi@lenovo.com; vbiradar@lenovo.com; dmishra6@lenovo.com; spinto@lenovo.com;
shanish@lenovo.com; rmallshetty@lenovo.com; vpandey6@lenovo.com; ganand3@lenovo.com;
ddyamangouda@lenovo.com; mroushan1@lenovo.com; rloitongbam@lenovo.com; kkanishka@lenovo.com;
nganachari@lenovo.com; apriya1@lenovo.com; aswami@lenovo.com; spanda7@lenovo.com;
jraj3@lenovo.com; mkaif@lenovo.com; ssahoo9@lenovo.com; fahmed6@lenovo.com;
sab@lenovo.com; tahamed2@lenovo.com; phanje@lenovo.com; agupta52@lenovo.com;
mhussain9@lenovo.com; ppriya4@lenovo.com; nkumar46@lenovo.com; cramani@lenovo.com;
aafrid@lenovo.com; ajoseph@lenovo.com; afernandez8@lenovo.com; dghosh@lenovo.com;
rsingh30@lenovo.com; slal@lenovo.com; smenon4@lenovo.com; sarbaz@lenovo.com;
vbhovi1@lenovo.com; mkr@lenovo.com
"""

DEFAULT_CC = "asinha10@lenovo.com; rsingh11@lenovo.com; mpintoo@lenovo.com"


def clean_recipients(text):
    """Accept addresses separated by ; , or new lines; return 'a; b; c'."""

    parts = text.replace(",", ";").replace("\n", ";").split(";")
    seen, result = set(), []
    for part in parts:
        address = part.strip()
        if address and address.lower() not in seen:
            seen.add(address.lower())
            result.append(address)
    return "; ".join(result)


# ============================================================
# AGEING BUCKETS
# ============================================================
# (label, min days, max days). Ageing = report date - created date in days.
# Emails created today (0 days) count in "01 Day". None = no upper limit.

AGEING_BUCKETS = [
    ("01 Day", 0, 1),
    ("02 Days", 2, 2),
    ("03 Days", 3, 3),
    ("04-05 Days", 4, 5),
    ("06-10 Days", 6, 10),
    ("11-20 Days", 11, 20),
    ("21-30 Days", 21, 30),
    ("31-60 Days", 31, 60),
    ("60+ Days", 61, None),
]

BUCKET_LABELS = [label for label, _, _ in AGEING_BUCKETS]

REQUIRED_COLUMNS = ["Current Sender (Object) (Email)", "Created On"]


def ageing_bucket(age_days):
    for label, low, high in AGEING_BUCKETS:
        if age_days >= low and (high is None or age_days <= high):
            return label
    return BUCKET_LABELS[0]


def ordinal(day):
    if 11 <= day % 100 <= 13:
        return f"{day}th"
    return f"{day}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th') }"


# ============================================================
# DATA PREPARATION
# ============================================================

def prepare_data(df, report_date):
    """Add Agent, Team Lead, ageing days and ageing bucket to the raw data."""

    df = df.copy()

    sender = df["Current Sender (Object) (Email)"].fillna("").astype(str).str.strip()
    key = sender.str.lower()

    df["Agent"] = key.map(canonical_name).fillna(sender)
    df["Mapped Team Lead"] = key.map(hc_lower)

    created = pd.to_datetime(df["Created On"], errors="coerce")
    if getattr(created.dt, "tz", None) is not None:
        created = created.dt.tz_localize(None)

    df["Created_Date"] = created.dt.normalize()
    df["Ageing (Days)"] = (
        (pd.Timestamp(report_date) - df["Created_Date"])
        .dt.days
        .fillna(0)
        .clip(lower=0)
        .astype(int)
    )
    df["Ageing Bucket"] = df["Ageing (Days)"].apply(ageing_bucket)

    return df


def bucket_counts(data):
    counts = data["Ageing Bucket"].value_counts()
    return {label: int(counts.get(label, 0)) for label in BUCKET_LABELS}


def build_reports(df):
    """Return (mapped records, TL summary, TL & Agent report)."""

    mapped = df[df["Mapped Team Lead"].notna()].copy()
    mapped["Team Lead"] = mapped["Mapped Team Lead"]

    tl_summary = (
        mapped.groupby("Team Lead")
        .size()
        .reset_index(name="Count of Team Lead")
        .sort_values(["Count of Team Lead", "Team Lead"], ascending=[False, True])
        .reset_index(drop=True)
    )

    rows = []

    for tl in tl_summary["Team Lead"]:
        tl_data = mapped[mapped["Team Lead"] == tl]

        rows.append(
            {"Team Lead": tl, "Agent": "", "Row Type": "TL",
             **bucket_counts(tl_data), "Grand Total": len(tl_data)}
        )

        agents = (
            tl_data.groupby("Agent")
            .size()
            .reset_index(name="Total")
            .sort_values(["Total", "Agent"], ascending=[False, True])
        )

        for agent, total in zip(agents["Agent"], agents["Total"]):
            agent_data = tl_data[tl_data["Agent"] == agent]
            rows.append(
                {"Team Lead": "", "Agent": agent, "Row Type": "Agent",
                 **bucket_counts(agent_data), "Grand Total": int(total)}
            )

    rows.append(
        {"Team Lead": "Grand Total", "Agent": "", "Row Type": "Grand Total",
         **bucket_counts(mapped), "Grand Total": len(mapped)}
    )

    columns = ["Team Lead", "Agent", "Row Type"] + BUCKET_LABELS + ["Grand Total"]
    report = pd.DataFrame(rows, columns=columns)

    return mapped, tl_summary, report


def validate(mapped, tl_summary, report):
    """Every total must agree; returns an error message or None."""

    tl_total = int(tl_summary["Count of Team Lead"].sum())
    agent_total = int(report.loc[report["Row Type"] == "Agent", "Grand Total"].sum())
    bucket_total = int(report.loc[report["Row Type"] == "Agent", BUCKET_LABELS].sum().sum())
    raw_total = len(mapped)

    if tl_total == agent_total == bucket_total == raw_total:
        return None

    return (
        "Calculation validation failed. "
        f"TL Total={tl_total}, Agent Total={agent_total}, "
        f"Bucket Total={bucket_total}, Mapped Raw Records={raw_total}"
    )


# ============================================================
# EXCEL
# ============================================================

HEADER_FONT = Font(name="Arial", size=10, bold=True, color="#f2f2f2")
NORMAL_FONT = Font(name="Arial", size=10)
BOLD_FONT = Font(name="Arial", size=10, bold=True, color="173B5A")
GRAND_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
TOTAL_FONT = Font(name="Arial", size=10, bold=True)

HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
TL_FILL = PatternFill("solid", start_color="D9EAF7", end_color="D9EAF7")
TOTAL_COL_FILL = PatternFill("solid", start_color="EAF2F8", end_color="EAF2F8")
GRAND_FILL = HEADER_FILL

THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Thousands separator, zeros shown as blank (same as the email).
COUNT_FORMAT = '#,##0;-#,##0;""'


def style_header(ws):
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = HEADER_ALIGN
    ws.row_dimensions[1].height = 30


def build_excel(raw_df, tl_summary, report):
    tl_export = pd.concat(
        [
            tl_summary,
            pd.DataFrame(
                {"Team Lead": ["Grand Total"],
                 "Count of Team Lead": [int(tl_summary["Count of Team Lead"].sum())]}
            ),
        ],
        ignore_index=True,
    )

    agent_export = report.drop(columns=["Row Type"])

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        tl_export.to_excel(writer, sheet_name="TL Wise", index=False)
        agent_export.to_excel(writer, sheet_name="TL & Agent Wise", index=False)
        raw_df.to_excel(writer, sheet_name="Raw Data", index=False)

        wb = writer.book

        # ---------------- TL Wise ----------------
        ws = wb["TL Wise"]
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A2"
        style_header(ws)

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            is_grand = row[0].value == "Grand Total"
            for cell in row:
                cell.border = BORDER
                cell.font = GRAND_FONT if is_grand else NORMAL_FONT
                if is_grand:
                    cell.fill = GRAND_FILL
                cell.alignment = LEFT if cell.column == 1 else CENTER
                if cell.column == 2:
                    cell.number_format = "#,##0"

        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 22

        # ------------- TL & Agent Wise -------------
        ws = wb["TL & Agent Wise"]
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "C2"
        style_header(ws)

        last_col = ws.max_column

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            team_lead = row[0].value
            is_grand = team_lead == "Grand Total"
            is_tl = bool(team_lead) and not row[1].value and not is_grand

            for cell in row:
                cell.border = BORDER

                if cell.column <= 2:
                    cell.alignment = LEFT
                else:
                    cell.alignment = CENTER
                    cell.number_format = COUNT_FORMAT

                if is_grand:
                    cell.font = GRAND_FONT
                    cell.fill = GRAND_FILL
                elif is_tl:
                    cell.font = BOLD_FONT
                    cell.fill = TL_FILL
                else:
                    cell.font = NORMAL_FONT
                    if cell.column == last_col:
                        cell.fill = TOTAL_COL_FILL
                        cell.font = TOTAL_FONT

            row[last_col - 1].number_format = "#,##0"

        ws.column_dimensions["A"].width = 36
        ws.column_dimensions["B"].width = 32
        for col in range(3, last_col):
            ws.column_dimensions[get_column_letter(col)].width = 12
        ws.column_dimensions[get_column_letter(last_col)].width = 13

        # ---------------- Raw Data ----------------
        ws = wb["Raw Data"]
        ws.sheet_view.showGridLines = False
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        style_header(ws)

        # Only date columns get a format; styling every raw cell is very slow
        # on 20k+ rows and adds nothing useful.
        for col_idx, name in enumerate(raw_df.columns, start=1):
            if pd.api.types.is_datetime64_any_dtype(raw_df[name]):
                date_format = "dd-mmm-yyyy" if name == "Created_Date" else "dd-mmm-yyyy hh:mm"
                for (cell,) in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                            min_col=col_idx, max_col=col_idx):
                    cell.number_format = date_format

        # Width from the header and the first 300 rows (fast on big files).
        for col_cells in ws.iter_cols(min_row=1, max_row=min(ws.max_row, 300)):
            longest = max(
                (len(str(c.value)) for c in col_cells if c.value is not None),
                default=0,
            )
            ws.column_dimensions[get_column_letter(col_cells[0].column)].width = (
                min(max(longest + 3, 12), 45)
            )

    return buffer.getvalue()


# ============================================================
# EMAIL STYLES
# ============================================================
# All styling is inline with fixed column widths, so the table keeps the same
# shape in new Outlook, Outlook on the web and classic Outlook.

EMAIL_STYLES = {
    "Corporate Blue": {
        "description": "Navy banner, Team Lead and Agent in separate columns, "
                       "zebra rows. Closest to the current report.",
        "banner": True,
        "kpis": False,
        "heat": False,
        "merged_names": False,
        "banner_bg": "#1F4E78",
        "header_bg": "#1F4E78",
        "header_fg": "#FFFFFF",
        "header_border": "#173B5A",
        "border": "#C9D6E2",
        "tl_bg": "#DDEBF7",
        "tl_fg": "#173B5A",
        "stripe": "#F5F8FB",
        "total_bg": "#EAF2F8",
        "tl_total_bg": "#CFE2F1",
        "grand_bg": "#1F4E78",
        "grand_fg": "#FFFFFF",
        "section_fg": "#1F4E78",
        "section_bar": "#2C7A9B",
        "section_bg": "#F1F6FA",
    },
    "Minimal Grey": {
        "description": "No banner, light grey headers, one compact name column. "
                       "Looks like a clean Excel table pasted into a normal email.",
        "banner": False,
        "kpis": False,
        "heat": False,
        "merged_names": True,
        "banner_bg": "#404040",
        "header_bg": "#F2F2F2",
        "header_fg": "#262626",
        "header_border": "#BFBFBF",
        "border": "#D9D9D9",
        "tl_bg": "#F7F7F7",
        "tl_fg": "#000000",
        "stripe": "#FFFFFF",
        "total_bg": "#FAFAFA",
        "tl_total_bg": "#EDEDED",
        "grand_bg": "#E7E6E6",
        "grand_fg": "#000000",
        "section_fg": "#262626",
        "section_bar": "#7F7F7F",
        "section_bg": "#FFFFFF",
    },
    "Priority Heatmap": {
        "description": "Summary tiles at the top and age columns coloured green to "
                       "red, so old pending emails stand out immediately.",
        "banner": True,
        "kpis": True,
        "heat": True,
        "merged_names": True,
        "banner_bg": "#1F3A5F",
        "header_bg": "#1F3A5F",
        "header_fg": "#FFFFFF",
        "header_border": "#16304F",
        "border": "#D5DEE7",
        "tl_bg": "#E9EEF5",
        "tl_fg": "#1F3A5F",
        "stripe": "#FFFFFF",
        "total_bg": "#F1F4F8",
        "tl_total_bg": "#D9E2EE",
        "grand_bg": "#1F3A5F",
        "grand_fg": "#FFFFFF",
        "section_fg": "#1F3A5F",
        "section_bar": "#1F3A5F",
        "section_bg": "#F4F6F9",
    },
}

DEFAULT_STYLE = "Corporate Blue"

# Heatmap colours per bucket: (header colour, light cell tint).
HEAT_COLORS = {
    "01 Day": ("#2E7D32", "#E8F5E9"),
    "02 Days": ("#2E7D32", "#E8F5E9"),
    "03 Days": ("#558B2F", "#F1F8E9"),
    "04-05 Days": ("#B7791F", "#FFF8E1"),
    "06-10 Days": ("#C05621", "#FFF0E0"),
    "11-20 Days": ("#C53030", "#FDECEA"),
    "21-30 Days": ("#9B2C2C", "#FADBD8"),
    "31-60 Days": ("#742A2A", "#F5C6C2"),
    "60+ Days": ("#4A1414", "#EFB3AD"),
}

KPI_GROUPS = [
    ("0-3 Days", ["01 Day", "02 Days", "03 Days"], "#2E7D32"),
    ("4-10 Days", ["04-05 Days", "06-10 Days"], "#C05621"),
    ("11-30 Days", ["11-20 Days", "21-30 Days"], "#C53030"),
    ("31+ Days", ["31-60 Days", "60+ Days"], "#742A2A"),
]

FONT = "font-family:Calibri,'Segoe UI',Arial,sans-serif;"

# Fixed column widths (px) keep every row and column the same size in Outlook.
W_TL, W_AGENT, W_NAME, W_BUCKET, W_TOTAL = 170, 180, 230, 54, 62

SIGNATURE_HTML = (
    "Regards,<br><br>"
    "<b>Preeti Barik</b><br>"
    "Escalation SPOC &ndash; Technical Support Customer Engagement Center<br>"
    "Lenovo India Pvt Ltd<br>"
    "Escalation Level 2: Sreejith Menon - "
    '<a href="mailto:smenon4@lenovo.com" style="color:#1F4E78;">smenon4@lenovo.com</a>'
)


def _table(width):
    return (
        f'<table cellpadding="0" cellspacing="0" border="0" width="{width}" '
        f'style="width:{width}px;border-collapse:collapse;table-layout:fixed;'
        f'mso-table-lspace:0pt;mso-table-rspace:0pt;{FONT}font-size:12px;'
        'margin:4px 0 18px 0;">'
    )


def _th(text, width, theme, bg=None):
    bg = bg or theme["header_bg"]

    return (
        f'<th width="{width}" bgcolor="{bg}" '
        f'style="width:{width}px;{FONT}font-size:11px;'
        'font-weight:bold;'
        'color:#FFFFFF !important;'
        f'background-color:{bg};'
        f'border:1px solid {theme["header_border"]};'
        'padding:5px 4px;'
        'text-align:center;'
        'vertical-align:middle;'
        'line-height:14px;'
        'mso-line-height-rule:exactly;">'
        '<span style="color:#FFFFFF !important;font-weight:bold;">'
        f'{text}'
        '</span>'
        '</th>'
    )

def _td(text, theme, bg, color="#262626", bold=False, align="center", indent=6):
    weight = "bold" if bold else "normal"
    return (
        f'<td bgcolor="{bg}" style="{FONT}font-size:12px;background-color:{bg};'
        f'color:{color};font-weight:{weight};text-align:{align};'
        f'border:1px solid {theme["border"]};padding:4px 6px 4px {indent}px;'
        'vertical-align:middle;line-height:16px;mso-line-height-rule:exactly;'
        'white-space:nowrap;overflow:hidden;">'
        f"{text}</td>"
    )


def _section(text, theme):
    return (
        '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;'
        'margin:16px 0 4px 0;"><tr>'
        f'<td bgcolor="{theme["section_bg"]}" style="border-left:3px solid '
        f'{theme["section_bar"]};background-color:{theme["section_bg"]};padding:5px 12px;'
        f'{FONT}font-size:14px;font-weight:bold;color:{theme["section_fg"]};">'
        f"{text}</td></tr></table>"
    )


def _paragraph(text):
    return f'<p style="{FONT}font-size:14px;color:#262626;margin:0 0 10px 0;">{text}</p>'


def _header_label(label):
    # "04-05 Days" -> "04-05<br>Days" so bucket columns can stay narrow.
    number, _, word = label.partition(" ")
    return f"{escape(number)}<br>{escape(word)}"


def _kpi_tiles(report, grand_total):
    totals = report[report["Row Type"] == "Grand Total"].iloc[0]
    tiles = [("Total Pending", grand_total, "#1F3A5F")]
    for name, labels, color in KPI_GROUPS:
        tiles.append((name, int(sum(int(totals[label]) for label in labels)), color))

    cells = "".join(
        f'<td width="128" bgcolor="{color}" style="width:128px;background-color:{color};'
        f'padding:10px 8px;text-align:center;border:3px solid #FFFFFF;">'
        f'<div style="{FONT}font-size:22px;font-weight:bold;color:#FFFFFF;'
        f'line-height:26px;">{value:,}</div>'
        f'<div style="{FONT}font-size:11px;color:#FFFFFF;line-height:14px;">{name}</div>'
        "</td>"
        for name, value, color in tiles
    )
    return (
        '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;'
        f'margin:4px 0 6px 0;"><tr>{cells}</tr></table>'
    )


def build_email_body(report_date, tl_summary, report, grand_total, style=DEFAULT_STYLE):
    """HTML body for Outlook (no <html> wrapper, no scripts)."""

    theme = EMAIL_STYLES[style]
    parts = []

    if theme["banner"]:
        parts.append(
            '<table cellpadding="0" cellspacing="0" border="0" width="100%" '
            'style="border-collapse:collapse;margin:0 0 16px 0;"><tr>'
            f'<td bgcolor="{theme["banner_bg"]}" style="background-color:{theme["banner_bg"]};'
            'padding:14px 18px;">'
            f'<div style="{FONT}font-size:20px;font-weight:bold;color:#FFFFFF;">'
            "Pending 2-Way Emails &ndash; Ageing Report</div>"
            f'<div style="{FONT}font-size:12px;color:#E7F0F8;margin-top:3px;">'
            f"Lenovo Customer Engagement Center &nbsp;|&nbsp; "
            f"{report_date.strftime('%d %B %Y')}</div>"
            "</td></tr></table>"
        )

    parts.append(_paragraph("Dear Associates,"))
    parts.append(_paragraph(
        "Kindly find the below pending 2-way emails as of today. "
        "Request you to review and clear them at the earliest."
    ))
    parts.append(_paragraph("Dear TLs,"))
    parts.append(_paragraph(
        "Kindly review the below details and support in getting these cleared on priority."
    ))

    if theme["kpis"]:
        parts.append(_kpi_tiles(report, grand_total))
    else:
        parts.append(_paragraph(
            f"<b>Report Date:</b> {report_date.strftime('%d %B %Y')}"
            "&nbsp;&nbsp;|&nbsp;&nbsp;"
            f"<b>Total Pending 2-Way Emails:</b> {grand_total:,}"
        ))

    # ---------------- Lead wise ----------------
    parts.append(_section("Lead Wise:", theme))
    parts.append(_table(W_TL + 130))
    parts.append(
        "<tr>" + _th("Team Lead", W_TL, theme) + _th("Count of Team Lead", 130, theme) + "</tr>"
    )

    for i, (tl, count) in enumerate(
        zip(tl_summary["Team Lead"], tl_summary["Count of Team Lead"])
    ):
        bg = "#FFFFFF" if i % 2 == 0 else theme["stripe"]
        parts.append(
            "<tr>"
            + _td(escape(str(tl)), theme, bg, bold=True, align="left", indent=8)
            + _td(f"{int(count):,}", theme, bg, bold=True)
            + "</tr>"
        )

    parts.append(
        "<tr>"
        + _td("Grand Total", theme, theme["grand_bg"], theme["grand_fg"], True, "left", 8)
        + _td(f"{grand_total:,}", theme, theme["grand_bg"], theme["grand_fg"], True)
        + "</tr></table>"
    )

    # ---------------- Agent wise ageing ----------------
    parts.append(_section("Agent Wise Count &amp; Ageing Overview:", theme))

    merged = theme["merged_names"]
    name_width = W_NAME if merged else W_TL + W_AGENT
    parts.append(_table(name_width + W_BUCKET * len(BUCKET_LABELS) + W_TOTAL))

    header = ["<tr>"]
    if merged:
        header.append(_th("Team Lead / Agent", W_NAME, theme))
    else:
        header.append(_th("Team Lead", W_TL, theme))
        header.append(_th("Agent", W_AGENT, theme))
    for label in BUCKET_LABELS:
        bg = HEAT_COLORS[label][0] if theme["heat"] else None
        header.append(_th(_header_label(label), W_BUCKET, theme, bg))
    header.append(_th("Grand<br>Total", W_TOTAL, theme))
    header.append("</tr>")
    parts.append("".join(header))

    agent_index = 0

    for _, row in report.iterrows():
        row_type = row["Row Type"]

        if row_type == "TL":
            bg, color, bold, total_bg = theme["tl_bg"], theme["tl_fg"], True, theme["tl_total_bg"]
        elif row_type == "Grand Total":
            bg, color, bold = theme["grand_bg"], theme["grand_fg"], True
            total_bg = theme["grand_bg"]
        else:
            bg = "#FFFFFF" if agent_index % 2 == 0 else theme["stripe"]
            color, bold, total_bg = "#262626", False, theme["total_bg"]
            agent_index += 1

        cells = []

        if merged:
            if row_type == "Agent":
                cells.append(_td(escape(str(row["Agent"])), theme, bg, color, bold, "left", 22))
            else:
                cells.append(_td(escape(str(row["Team Lead"])), theme, bg, color, bold, "left", 8))
        else:
            if row_type == "Agent":
                cells.append(_td("&nbsp;", theme, bg, color, bold, "left", 8))
                cells.append(_td(escape(str(row["Agent"])), theme, bg, color, bold, "left", 8))
            else:
                cells.append(_td(escape(str(row["Team Lead"])), theme, bg, color, bold, "left", 8))
                cells.append(_td("&nbsp;", theme, bg, color, bold, "left", 8))

        for label in BUCKET_LABELS:
            value = int(row[label])
            cell_bg = bg
            if theme["heat"] and value and row_type == "Agent":
                cell_bg = HEAT_COLORS[label][1]
            cells.append(_td(f"{value:,}" if value else "&nbsp;", theme, cell_bg, color, bold))

        cells.append(_td(f"{int(row['Grand Total']):,}", theme, total_bg, color, True))

        parts.append("<tr>" + "".join(cells) + "</tr>")

    parts.append("</table>")

    # ---------------- Signature ----------------
    parts.append(
        f'<p style="{FONT}font-size:14px;color:#262626;margin:18px 0 0 0;">'
        f"{SIGNATURE_HTML}</p>"
    )

    return f'<div style="{FONT}font-size:14px;color:#262626;">' + "".join(parts) + "</div>"


def outlook_html(body):
    return (
        '<!DOCTYPE html><html><head><meta charset="UTF-8">'
        "</head><body>" + body + "</body></html>"
    )


def preview_html(body):
    """Streamlit preview with a one-click copy button."""

    return f"""<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:12px;background:#F3F6FA;">
<div style="background:#FFFFFF;border:1px solid #DCE4EC;padding:20px;">
<div id="email-content">{body}</div>
</div>
<br>
<button onclick="copyEmailContent()" style="background-color:#1F4E78;color:#FFFFFF;
padding:11px 20px;border:none;border-radius:5px;font-weight:bold;cursor:pointer;font-size:14px;">
📋 Copy Formatted Email to Clipboard</button>
<span id="copy-status" style="margin-left:10px;font-family:Arial;font-size:13px;
color:green;font-weight:bold;"></span>
<script>
function copyEmailContent() {{
  var range = document.createRange();
  range.selectNode(document.getElementById("email-content"));
  var selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
  var status = document.getElementById("copy-status");
  try {{
    status.innerText = document.execCommand("copy")
      ? "Copied! Tip: the 'Create Draft' button keeps formatting better than pasting."
      : "Copy failed.";
  }} catch (err) {{
    status.innerText = "Copy failed.";
  }}
  selection.removeAllRanges();
}}
</script>
</body></html>"""


# ============================================================
# OUTLOOK DRAFT
# ============================================================
# The new Outlook cannot be automated from a script, so the draft is written
# into your Exchange mailbox through the Outlook desktop engine (it stays in the
# background, no window opens). Exchange syncs it, so it shows up in the
# Drafts folder of the new Outlook within a few seconds.

OUTLOOK_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$outlook = New-Object -ComObject Outlook.Application
$session = $outlook.Session
$mail = $outlook.CreateItem(0)
if ($env:DRAFT_TO) { $mail.To = $env:DRAFT_TO }
if ($env:DRAFT_CC) { $mail.CC = $env:DRAFT_CC }
$mail.Subject = $env:DRAFT_SUBJECT
$mail.HTMLBody = [IO.File]::ReadAllText($env:DRAFT_BODY_FILE, [Text.Encoding]::UTF8)
[void]$mail.Attachments.Add($env:DRAFT_ATTACHMENT_FILE)
[void]$mail.Recipients.ResolveAll()
$mail.Save()
try { $session.SendAndReceive($false) } catch { }
Start-Sleep -Seconds 10
Write-Output "SAVED"
"""


# Outlook automation only exists on Windows. On Streamlit Cloud (Linux) the app
# offers a ready-made draft file (.eml) instead.
CAN_AUTOMATE_OUTLOOK = os.name == "nt"


def build_eml(subject, html_body, attachment_name, attachment_data, to="", cc=""):
    """Unsent email file: opens in Outlook as a draft ready to send."""

    msg = EmailMessage()
    msg["Subject"] = subject
    if to:
        msg["To"] = ", ".join(a.strip() for a in to.split(";") if a.strip())
    if cc:
        msg["Cc"] = ", ".join(a.strip() for a in cc.split(";") if a.strip())
    msg["X-Unsent"] = "1"  # tells Outlook to open it as an editable draft

    plain = re.sub(r"<[^>]+>", " ", html_body)
    plain = re.sub(r"\s+", " ", plain.replace("&nbsp;", " ")).strip()
    msg.set_content(plain)
    msg.add_alternative(html_body, subtype="html")
    msg.add_attachment(
        attachment_data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment_name,
    )
    return msg.as_bytes()


def create_outlook_draft(subject, html_body, attachment_name, attachment_data, to="", cc=""):
    """Save an unsent draft (with attachment) in the signed-in Exchange mailbox."""

    with tempfile.TemporaryDirectory() as folder:
        body_file = os.path.join(folder, "body.html")
        attachment_file = os.path.join(folder, attachment_name)

        with open(body_file, "w", encoding="utf-8") as file:
            file.write(html_body)
        with open(attachment_file, "wb") as file:
            file.write(attachment_data)

        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", OUTLOOK_SCRIPT],
            env={
                **os.environ,
                "DRAFT_SUBJECT": subject,
                "DRAFT_TO": to,
                "DRAFT_CC": cc,
                "DRAFT_BODY_FILE": body_file,
                "DRAFT_ATTACHMENT_FILE": attachment_file,
            },
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

    if result.returncode != 0 or "SAVED" not in result.stdout:
        raise RuntimeError(result.stderr.strip() or "Outlook did not respond.")


# ============================================================
# STREAMLIT APP
# ============================================================

PAGE_CSS = """
<style>
.main-title {font-size:28px;font-weight:700;color:#1F4E78;margin-bottom:4px;}
.sub-text {font-size:14px;color:#666666;margin-bottom:20px;}
.stDownloadButton button, .stButton button {
    background-color:#1F4E78;color:white;font-weight:700;border-radius:6px;border:none;}
.stDownloadButton button:hover, .stButton button:hover {background-color:#163A5C;color:white;}
.metric-box {padding:14px;border-radius:8px;background-color:#F3F7FA;
    border-left:5px solid #1F4E78;margin-bottom:10px;}
.metric-title {font-size:12px;color:#666666;}
.metric-value {font-size:24px;font-weight:700;color:#1F4E78;}
</style>
"""


@st.cache_data(show_spinner="Building the report...")
def process_file(file_bytes, report_date):
    """Read, calculate and build the Excel once per uploaded file and day."""

    raw = pd.read_excel(BytesIO(file_bytes), sheet_name=0)

    missing = [c for c in REQUIRED_COLUMNS if c not in raw.columns]
    if missing:
        return {"missing": missing}

    df = prepare_data(raw, report_date)
    mapped, tl_summary, report = build_reports(df)

    return {
        "missing": [],
        "df": df,
        "mapped": mapped,
        "tl_summary": tl_summary,
        "report": report,
        "error": validate(mapped, tl_summary, report),
        "excel": build_excel(df, tl_summary, report),
    }


def metric(column, title, value):
    column.markdown(
        f'<div class="metric-box"><div class="metric-title">{title}</div>'
        f'<div class="metric-value">{value}</div></div>',
        unsafe_allow_html=True,
    )


def main():
    st.set_page_config(
        page_title="Lenovo CEC 2-Way Email Ageing Generator",
        page_icon="📧",
        layout="wide",
    )
    st.markdown(PAGE_CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="main-title">Lenovo CEC - 2-Way Email Ageing Report Generator</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="sub-text">Upload the latest raw Excel file to generate the TL-wise '
        "summary, TL &amp; Agent ageing report, formatted Excel workbook and an "
        "Outlook draft ready to send.</div>",
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader("Upload Raw Excel File (.xlsx)", type=["xlsx"])

    if uploaded_file is None:
        st.info("Please upload the latest raw 2-way email Excel file to generate the report.")
        st.stop()

    report_date = pd.Timestamp.today().normalize().date()

    try:
        result = process_file(uploaded_file.getvalue(), report_date)
    except Exception as error:
        st.error(f"Unable to read the Excel file: {error}")
        st.stop()

    if result["missing"]:
        st.error(
            "The uploaded Excel file is missing the following required "
            f"column(s): {', '.join(result['missing'])}"
        )
        st.stop()

    if result["error"]:
        st.error(f"⚠️ {result['error']}")
        st.stop()

    df = result["df"]
    mapped = result["mapped"]
    tl_summary = result["tl_summary"]
    report = result["report"]
    excel_data = result["excel"]

    grand_total = len(mapped)

    c1, c2, c3, c4 = st.columns(4)
    metric(c1, "Report Date", report_date.strftime("%d %b %Y"))
    metric(c2, "Total Raw Records", f"{len(df):,}")
    metric(c3, "Mapped Records", f"{grand_total:,}")
    metric(c4, "Team Leads", f"{len(tl_summary):,}")

    st.success(f"Calculation validated successfully: {grand_total:,} mapped records.")

    invalid_dates = int(df["Created_Date"].isna().sum())
    if invalid_dates:
        st.warning(
            f"{invalid_dates:,} record(s) have a blank or invalid 'Created On' date; "
            "they are counted in the '01 Day' bucket."
        )

    unmapped = df.loc[df["Mapped Team Lead"].isna(), "Agent"]
    if len(unmapped):
        with st.expander(
            f"ℹ️ {len(unmapped):,} record(s) not in the TL mapping (excluded from the report)"
        ):
            st.dataframe(
                unmapped.replace("", "(blank)").value_counts()
                .rename_axis("Agent").reset_index(name="Records"),
                hide_index=True,
            )

    # ---------------- Excel ----------------
    excel_name = f"Lenovo_CEC_2Way_Ageing_{report_date.strftime('%d_%b_%Y')}.xlsx"

    st.markdown("---")
    st.subheader("📊 Excel Report")
    st.download_button(
        label="📥 Download Formatted Excel Report",
        data=excel_data,
        file_name=excel_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ---------------- Email ----------------
    st.markdown("---")
    st.subheader("✉️ Email")

    style = st.radio(
        "Email body style",
        list(EMAIL_STYLES),
        index=list(EMAIL_STYLES).index(DEFAULT_STYLE),
        horizontal=True,
    )
    st.caption(EMAIL_STYLES[style]["description"])

    email_body = build_email_body(report_date, tl_summary, report, grand_total, style)

    subject_default = (
        f"Pending 2-Way Emails as of {ordinal(report_date.day)} "
        f"{report_date.strftime('%b %Y')}"
    )

    subject = st.text_input("Subject", value=subject_default)
    to = st.text_area("To", value=clean_recipients(DEFAULT_TO), height=110)
    cc = st.text_area("CC", value=DEFAULT_CC, height=68)

    eml_data = build_eml(
        subject=subject,
        html_body=outlook_html(email_body),
        attachment_name=excel_name,
        attachment_data=excel_data,
        to=clean_recipients(to),
        cc=clean_recipients(cc),
    )
    eml_name = f"Pending_2Way_Emails_{report_date.strftime('%d_%b_%Y')}.eml"

    if not CAN_AUTOMATE_OUTLOOK:
        st.info(
            "Download the draft below and double-click it: Outlook opens it as a new "
            "email with To, CC, subject, the formatted body and the Excel attached. "
            "Review it and click Send.\n\n"
            "If it opens as a read-only message instead, right-click the file → "
            "Open with → Outlook (classic)."
        )
        st.download_button(
            label="📥 Download Outlook Draft (.eml)",
            data=eml_data,
            file_name=eml_name,
            mime="message/rfc822",
        )

    elif st.button("📝 Create Draft in Outlook (To, CC and Excel included)"):
        try:
            with st.spinner("Saving the draft to your mailbox..."):
                create_outlook_draft(
                    subject=subject,
                    html_body=outlook_html(email_body),
                    attachment_name=excel_name,
                    attachment_data=excel_data,
                    to=clean_recipients(to),
                    cc=clean_recipients(cc),
                )
            st.success(
                "Draft saved. Open the new Outlook → Drafts to review and send it "
                "(it may take a few seconds to appear). Nothing was sent."
            )
        except (OSError, RuntimeError, subprocess.TimeoutExpired) as err:
            st.error(
                "Could not create the draft. The Outlook desktop app (classic) must be "
                "installed and signed in to your mailbox once; it does not need to be "
                f"open.\n\n{err}"
            )

    if st.checkbox("Show email preview", value=True):
        st.components.v1.html(preview_html(email_body), height=1100, scrolling=True)
        st.download_button(
            label="📥 Download Email as HTML",
            data=outlook_html(email_body),
            file_name=f"2_Way_Email_Ageing_{report_date.strftime('%d_%b_%Y')}.html",
            mime="text/html",
        )


if __name__ == "__main__":
    main()
