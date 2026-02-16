import dash
from dash import dcc, html, Input, Output, State, ctx, callback_context, ALL
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import plotly.graph_objects as go
import base64
import io
import json
import re
import string
import os
import io
import os
import re
import string
from datetime import datetime
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from dash import dcc, html  
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.dml.color import RGBColor
import plotly.io as pio
from io import BytesIO
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import dash
from dash import dcc, html, Input, Output, State
import pandas as pd
import numpy as np
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
import base64
import io
import pandas as pd
import plotly.express as px
import plotly.io as pio
import re
import string
import io
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches
from dash import dcc
from dash.dependencies import Input, Output, State
import os
from pptx import Presentation
import dash
from dash import dcc, html, Input, Output, State, ctx
import pandas as pd
import base64
import io
from datetime import datetime
from dash.dependencies import MATCH, ALL
import pandas as pd
import numpy as np
import re
from datetime import datetime
from dash import dcc, html, Input, Output, State
import plotly.express as px
import dash
from dash import dcc, html, Input, Output, State, ALL, callback_context
import pandas as pd
import base64
import io
import json
import dash
from dash import dcc, html, Input, Output, State, ALL, callback_context
import pandas as pd
import base64
import io
import json
import dash
from dash import dcc, html, Input, Output, State, ALL, callback_context
import pandas as pd
import base64
import io
import json
from dash import html, dcc, dash_table
from dash.dependencies import Input, Output, State
import pandas as pd
from datetime import datetime
from io import BytesIO
from datetime import datetime, date 
import webbrowser
from threading import Timer

# ============================================================
# APP CONFIG
# ============================================================

app = dash.Dash(__name__, assets_folder="assets", suppress_callback_exceptions=True)
app.title = "Incident Analysis Dashboard"
server = app.server

PRIMARY_COLOR = "#003135"
SECONDARY_COLOR = "#00e28b"
WHITE = "#ffffff"
BLACK = "#000000"




# ======================= 
# Issue Categorization Logic
#========================

groups = [
    {
        "name": "Disk Related Issues",
        "keywords": [
            "% free space", "disk usage", "disk space", "disk full", "win_disk_all_wmi",
            "win_disk_all_nrpe", "win_mountpoint_f_sqlbackup", "disk of", "disk minor",
            "disk warning", "disk critical", "disk"
        ]
    },
    {
        "name": "CPU Threshold Breaching",
        "keywords": [
            "cpu", "cpu threshold", "cpu usage", "cpu critical state", "win_cpu_utilization_nrpe",
            "cpu minor", "Threshold Breach CPUPercentage"
        ]
    },
    {
        "name": "SQL is in Unknown state",
        "keywords": [
            "SQL is in Unknown state"
        ]
    },
    {
        "name": "Memory Threshold Breaching",
        "keywords": [
            "mem availability", "Threshold Breach MEM Availability is less than 1GB",
            "memory low", "RAM utilization", "memory MINOR", "RAM Warning threshold",
            "win_mem_physical_nrpe","nix_mem_utilization","RAM is Warning","RAM","mssql_total-server-memory"
        ]
    },
    {
        "name": "NADC Blocking",
        "keywords": [
            "event_log_id_603", "NADC"
        ]
    },
    {
        "name": "Job Failures",
        "keywords": [
            "failed job", "winevent_id_208", "job failed", "SQL Job", "Event_208",
            "mssql_jobfailures_tds", "mssql_failed_jobs", "Windows Event ID 208"
        ]
    },
    {
        "name": "Backup Recovery & Restore",
        "keywords": [
            "Backup Recovery & Restore"
        ]
    },
    {
        "name": "Backup Releted",
        "keywords": [
            "failed backup", "MSSQL Failed Backup", "incorrect backup", "missing files",
            "backup", "fullbackup", "Alert: SQL Backup Failures Type: Job Management - Data Protection"
        ]
    },
        {
        "name": "Execute DDL Script on Production",
        "keywords": [
            "Run the provided DDL script on the production database"
        ]
    },
    {
        "name": "Database Migration from UAT to Production",
        "keywords": [
            "Master proposal:Copy DBS/IBS/LBS DPR Products database from UAT to Production"
        ]
    },
    {
        "name": "SQL Server Agent Job Management",
        "keywords": [
            "SQL Server Agent Job creation or modification","Create new SQL Agent Job"
        ]
    },
    {
        "name": "Cleanup Unused SQL Logins",
        "keywords": [
            "Delete the selected logins that are not required on the SQL instances on Woob burry servers","Delete the  selected logins"
        ]
    },
    {
        "name": "Service Account Password Reset",
        "keywords": [
            "Service Accounts - Password Reset on prod server"
        ]
    },
    {
        "name": "Enable Encryption on Non-Production SQL Servers",
        "keywords": [
            "Enable encryption","Enabling the encryption"
        ]
    },
        {
        "name": "win_service_ncpa1",
        "keywords": [
            "win_service_ncpa"
        ]
    },
    {
        "name": "Restore Releted",
        "keywords": [
            "SQL Database Restore Request", "Need to restore", "restore fail","Restore"
        ]
    },
    {
        "name": "SQL Service Down",
        "keywords": [
            "sql server availability", "service was not running", "component sql server stopped",
            "MSSQLSERVER has stopped", "SQLSERVERAGENT has stopped", "Service Down",
            "SQLSERVERAGENT", "SSRS down", "MSSQL SQL Server Availability",
            "MSSQL SQL ServerAgent Availability", "win_service_wmi_do" 
        ]
    },
    {
        "name": "Degraded Service",
        "keywords": [
            "Degraded Service"
        ]
    },
    {
        "name": "Critical Events",
        "keywords": [
            "MSSQL Critical Event", "critical event", "event_id_601", "windows event id 601",
            "event_log_id_601", "event_log_id_602", "event_14421", "new event_id",
            "event_log_id", "windows event id", "event_"
        ]
    },
    {
        "name": "Error replicating data",
        "keywords": [
            "Error replicating data"
        ]
    },
    {
        "name": "Databases with no owners",
        "keywords": [
            "Databases with no owners"
        ]
    },
    {
        "name": "LogicMonitor",
        "keywords": [
            "LogicMonitor"
        ]
    },
    {
        "name": "SQL ServiceNow Tickets",
        "keywords": [
            "SQL ServiceNow Tickets"
        ]
    },
    {
        "name": "Azure: Deactivated Severity",
        "keywords": [
            "Azure: Deactivated Severity"
        ]
    },
    {
        "name": "SAzure: Activated Severity",
        "keywords": [
            "Azure: Activated Severity"
        ]
    },
    {
        "name": "Refresh failed: Commvault Logs has failed to refresh",
        "keywords": [
            "Refresh failed: Commvault Logs has failed to refresh"
        ]
    },
    {
        "name": "New DB  access request - SQL Database Access Request",
        "keywords": [
            "New DB  access request - SQL Database Access Request"
        ]
    },
    {
        "name": "Job Change Details notification",
        "keywords": [
            "Job Change Details notification"
        ]
    },
    {
        "name": "PagerDuty overdue functions calls - See PagerDutyDC function app in Azure ",
        "keywords": [
            "PagerDuty", "PagerDuty overdue functions calls - See PagerDutyDC function app in Azure "
        ]
    },
    {
        "name": "Databases with no owners",
        "keywords": [
            "Databases with no owners"
        ]
    },
    {
        "name": "Access / Lock Issues",
        "keywords": [
            "lock", "block", "queryblock", "Blocking transactions", "misconfiguration",
            "no access", "unable to access", "lost access", "cannot access"
        ]
    },
    {
        "name": "Requesting Access / Permission",
        "keywords": [
            "Requesting access", "Need access", "access required", "creation request",
            "Request for Database Access", "Need user access"
        ]
    },
    {
        "name": "Data Load Failures",
        "keywords": [
            "data load failures", "load is failing", "job unable to write", "receiving data files late"
        ]
    },
    {
        "name": "SQL Installation Issues",
        "keywords": [
            "installation request", "can't find server", "sql server not installed",
            "sql installed is not working", "sql installation request"
        ]
    },
    {
        "name": "Database Maintenance Actions",
        "keywords": [
            "restart database"
        ]
    },
    {
        "name": "Connection / Login Issues",
        "keywords": [
            "unable to connect", "connect", "login failure", "login failed", "Not able to login",
            "SQL Server Login/Access Request", "error establishing connection", "connection timeout",
            "create the login for the server"
        ]
    },
    {
        "name": "Database Down / Unavailable",
        "keywords": [
            "database offline", "database unavailable", "mssql_database_online", "db bounce",
            "database down", "is down", "vm availability", "host down"
        ]
    },
    {
        "name": "Connection Problems",
        "keywords": [
            "conn major", "connection problem", "connection failed", "unable to connect",
            "sql - unable to see data", "mssql_connection_time", "mssql_connection_time1"
        ]
    },
    {
        "name": "Certificate / Licence Issues",
        "keywords": [
            "certificate", "Licence"
        ]
    },
    {
        "name": "Database script execution Request",
        "keywords": [
            "Database script execution Request", "Request to execute SQL script in PROD"
        ]
    },
    {
        "name": "MSBI / Reporting Failures",
        "keywords": [
            "msbi reports failing"
        ]
    },
    {
        "name": "Applying Patches / Upgrades",
        "keywords": [
            "patch", "Sql patch", "Apply latest patches", "Apply latest Cumulative update",
            "Apply latest", "Apply latest CU", "-->", " Apply latest SP"
        ]
    },
    {
        "name": "SQL Installation",
        "keywords": [
            "sql installation",
            "install sql",
            "sql server install",
            "new sql instance",
            "sql setup",
            "sql deployment",
            "sql build",
            "fresh sql installation",
            "sql server installation request"
        ]
    },
    {
        "name": "Decommission Tasks",
        "keywords": [
            "decommission tasks",
            "decommission database",
            "decommission",
            "sql decommission",
            "db decommission",
            "server decommission",
            "retire database",
            "remove sql instance",
            "decom request"
        ]
    },
    {
        "name": "Create Service Account",
        "keywords": [
            "create service account",
            "service account creation",
            "sql service account",
            "domain service account",
            "create ad account",
            "service account request",
            "managed service account",
            "msa creation"
        ]
    },
    {
        "name": "Azure Request - Set SQL Permissions",
        "keywords": [
            "azure request - set sql permissions",
            "set sql permissions",
            "azure sql permissions",
            "grant sql access",
            "azure sql role",
            "add user to sql",
            "sql permission request",
            "azure sql access",
            "grant db permissions"
        ]
    },


]
# =======================


# ============================================================
# INDEX STRING (FULL WEBSITE LOOK)
# ============================================================

app.index_string = """
<!DOCTYPE html>
<html>
<head>
    {%metas%}
    <title>{%title%}</title>
    {%favicon%}
    {%css%}

    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap" rel="stylesheet">

    <style>
        :root {
            --accent: #00e28b;
        }

        html, body {
            height: 100%;
            margin: 0;
        }

        body {
            font-family: 'Inter', sans-serif;
            background: radial-gradient(circle at top, #044a4f, #003135 70%);
            color: white;
        }

        /* ===== PAGE WRAPPER ===== */
        .page {
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            padding: 16px;
        }

        /* ===== MAIN CONTENT ===== */
        .container {
            max-width: 1400px;
            margin: auto;
            width: 100%;
            flex: 1;
        }

        .glass {
            background: rgba(255,255,255,0.06);
            backdrop-filter: blur(14px);
            border-radius: 18px;
            border: 1px solid rgba(255,255,255,0.15);
            padding: 24px;
            margin-bottom: 24px;
            box-shadow: 0 12px 30px rgba(0,0,0,0.35);
        }

        .title {
            font-size: clamp(28px, 5vw, 42px);
            font-weight: 700;
            background: linear-gradient(45deg, #00e28b, #00b8d4);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .upload-box {
            padding: 32px;
            border: 2px dashed var(--accent);
            border-radius: 20px;
            text-align: center;
            cursor: pointer;
        }

        /* ---------- Sheet Buttons ---------- */

        .sheet-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 16px;
            margin-top: 20px;
        }

        .sheet-btn {
            background: rgba(255,255,255,0.08);
            border: 2px solid rgba(0,226,139,0.4);
            color: white;
            padding: 18px;
            border-radius: 16px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.25s ease;
        }

        .sheet-btn:hover {
            transform: translateY(-3px);
        }

        .sheet-btn.selected {
            background: linear-gradient(45deg, #00e28b, #00b8d4);
            color: #002b2f;
            border-color: #00e28b;
        }

        /* ---------- Filters ---------- */

        .filter-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
            gap: 20px;
        }

        .input-box {
            padding: 14px;
            border-radius: 28px;
            border: 2px solid var(--accent);
            font-size: 16px;
            font-weight: 600;
            text-align: center;
            width: 95%;
        }
        .Select-menu-outer {
            background-color: #003135 !important;   /* PRIMARY_COLOR */
        }

        .Select-control {
            height: 64px !important;
            border-radius: 28px !important;
            border: 2px solid var(--accent) !important;
        }

        .Select-value {
            line-height: 64px !important;
            font-size: 18px !important;
            font-weight: 700 !important;
            text-align: center;
        }

        /* ---------- Buttons ---------- */

        .btn-primary, .btn-secondary {
            width: 100%;
            padding: 16px;
            border-radius: 28px;
            font-size: 18px;
            font-weight: 700;
            border: none;
            cursor: pointer;
            transition: transform 0.15s ease;
        }

        .btn-primary {
            background: linear-gradient(45deg, #00e28b, #00b8d4);
            color: #002b2f;
        }

        .btn-secondary {
            background: linear-gradient(45deg, #ff6b6b, #ff8e8e);
            color: white;
        }

        .btn-primary:active,
        .btn-secondary:active {
            transform: scale(0.96);
        }

        .button-row {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
            gap: 16px;
            margin-top: 24px;
        }

        .selected-sheet-badge {
            margin-top: 12px;
            font-weight: 600;
            color: #00e28b;
            text-align: center;
        }

        /* =====================================================
           CHART SECTION (IMPROVED & POLISHED)
           ===================================================== */

        .chart-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(420px, 1fr));
            gap: 28px;
            margin-top: 28px;
            align-items: stretch;
        }

        .chart-card {
            background: rgba(255,255,255,0.06);
            backdrop-filter: blur(14px);
            border-radius: 20px;
            border: 1px solid rgba(255,255,255,0.18);
            padding: 24px;
            box-shadow: 0 14px 35px rgba(0,0,0,0.35);
            transition: transform 0.25s ease, box-shadow 0.25s ease;
        }

        .chart-card:hover {
            transform: translateY(-6px);
            box-shadow: 0 22px 45px rgba(0,0,0,0.45);
        }

        .chart-card h4 {
            margin-top: 0;
            margin-bottom: 6px;
            letter-spacing: 0.5px;
        }

        .chart-card p {
            margin-top: 0;
            margin-bottom: 16px;
            opacity: 0.85;
        }

        @media (max-width: 1200px) {
            .chart-grid {
                grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
            }
        }

        @media (max-width: 768px) {
            .chart-grid {
                grid-template-columns: 1fr;
                gap: 20px;
            }

            .chart-card {
                padding: 20px;
            }
        }

        







        /* ===== P1 TICKETS (AUTO FLOW / N-COLUMN WRAP) ===== */

        .p1-section {
            margin-top: 24px;
        }

        .p1-title {
            text-align: center;
            color: #ff6b6b;
            font-weight: 700;
            margin-bottom: 12px;
        }

        /* 🔥 AUTO-WRAPPING CONTAINER (N COLUMNS) */
        .p1-grid {
            display: flex;
            flex-wrap: wrap;
            justify-content: center;
            gap: 10px 16px;
        }

        /* 🔥 INDIVIDUAL TICKET CHIP */
        .p1-item {
            font-size: 0.9rem;
            font-weight: 600;
            color: #ffd1d1;
            background: rgba(255, 107, 107, 0.15);
            border: 1px solid rgba(255, 107, 107, 0.45);
            padding: 6px 14px;
            border-radius: 12px;
            white-space: nowrap;
            min-width: 90px;
            text-align: center;
        }








        /* ---------- FOOTER ---------- */

        .footer {
            background: rgba(0,0,0,0.35);
            backdrop-filter: blur(10px);
            border-top: 1px solid rgba(255,255,255,0.15);
            padding: 14px 20px;
            text-align: center;
            font-size: 14px;
            color: #d0f5ea;
        }

        .footer strong {
            color: #00e28b;
        }
    </style>
</head>

<body>
    {%app_entry%}
    <footer>
        {%config%}
        {%scripts%}
        {%renderer%}
    </footer>
</body>
</html>
"""

# ============================================================
# LAYOUT
# ============================================================

app.layout = html.Div(className="page", children=[

    html.Div(className="container", children=[

        # ================= TITLE =================
        html.Div(className="glass", children=[html.Div("Incident Analysis Dashboard", className="title")]),

        # ================= UPLOAD =================
        html.Div(className="glass", children=[
            dcc.Upload(id="upload-data", className="upload-box", children="Drag & Drop or Click to Upload Excel", multiple=False),
            html.Div(id="upload-output", style={"marginTop": "12px"})
        ]),

        html.Div(id="analyze-section"),
        html.Div(id="filter-controls"),

        # ================= CHARTS WRAPPER =================
        html.Div(id="charts-wrapper", style={"display": "none"}, children=[

            # KPI SUMMARY
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="total-tickets-summary"))),
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color="red", children=html.Div(id="usergen-incident-summary"))),
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color="red", children=html.Div(id="unassigned-incident-summary")))
            ]),
            #SLA SUMMERY
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-grid", children=[html.Div(className="chart-card", children=[dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="assigned-unassigned-gauge"))])])

            ]),

            # COMPANY & INSTANCE
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-card", children=dcc.Loading(id="company-chart-loading", type="circle", color=SECONDARY_COLOR, children=html.Div(id="company-bar-chart"))),
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="tickets-per-instance-table")))
            ]),

            # MONTH TREND
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="month-bar-chart")))
            ]),
            html.Div(
                className="chart-grid",
                children=[
                    html.Div(
                        className="chart-card",
                        children=dcc.Loading(
                            type="circle",
                            color=SECONDARY_COLOR,
                            children=html.Div(id="month-clustered-chart")
                        )
                    )
                ]
            ),

            # STATE & PRIORITY
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="state-pie-chart"))),
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="priority-bar-chart")))
            ]),

            # CONFIG & TOP PROBLEMS
            # CONFIG & TOP PROBLEMS
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-card", children=[dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="configitem-bar-chart")), 
                html.Div(dcc.Dropdown(id="config-item-dropdown", placeholder="Select Configuration Item", clearable=False, searchable=True), id="config-item-container", style={"display": "none"})]), 
                html.Div(className="chart-card", children=[dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="top-problems-chart1")), dcc.Download(id="download-problems-xlsx")])
            ]),

            # TIME OF DAY
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="time-of-day-distribution")))
            ]),

            # ASSIGNMENT GROUP & TYPE
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="assignment-group-bar-chart"))),
                html.Div(className="chart-card", children=dcc.Loading(id="type-chart-loading", type="circle", color=SECONDARY_COLOR, children=html.Div(id="type-bar-chart")))
            ]),

            # ASSIGNED TO
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="assigned-to-bar-chart")))
            ]),

            # RESOLVED BY
            html.Div(className="chart-grid", children=[
                html.Div(className="chart-card", children=dcc.Loading(type="circle", color=SECONDARY_COLOR, children=html.Div(id="resolved-by-bar-chart")))
            ])

        ]),

        # ================= STORES =================
        dcc.Store(id="stored-file"),
        dcc.Store(id="stored-df"),
        dcc.Store(id="selected-sheet-name"),  # ⭐ NEW STORE FOR SHEET NAME

    ]),

    # ================= FOOTER =================
    html.Div(className="footer", children=f"© {datetime.now().year} Incident Analysis Dashboard | Powered by Unisys Database Solutions | SRM")

])

# ============================================================
# UPLOAD CALLBACK
# ============================================================

@app.callback(
    [Output("upload-output", "children"),
     Output("stored-file", "data"),
     Output("analyze-section", "children")],
    Input("upload-data", "contents"),
    State("upload-data", "filename"),
    prevent_initial_call=True
)
def handle_upload(contents, filename):
    content_string = contents.split(",")[1]
    decoded = base64.b64decode(content_string)

    xls = pd.ExcelFile(io.BytesIO(decoded))

    buttons = [
        html.Button(
            sheet,
            id={"type": "sheet-button", "index": i},
            n_clicks=0,
            className="sheet-btn"
        )
        for i, sheet in enumerate(xls.sheet_names)
    ]

    return (
        f"Uploaded: {filename}",
        content_string,
        html.Div(className="glass", children=[
            html.H3("Select Sheet"),
            html.Div(buttons, className="sheet-grid")
        ])
    )


# ============================================================
# SHEET SELECTION → STORE SHEET NAME STRING
# ============================================================

@app.callback(
    [
        Output("filter-controls", "children"),
        Output("stored-df", "data"),
        Output("selected-sheet-name", "data"),   # ⭐ HERE
        Output({"type": "sheet-button", "index": ALL}, "className")
    ],
    Input({"type": "sheet-button", "index": ALL}, "n_clicks"),
    State("stored-file", "data"),
    prevent_initial_call=True
)
def load_sheet(n_clicks, file_data):
    ctx = callback_context
    idx = json.loads(ctx.triggered[0]["prop_id"].split(".")[0])["index"]

    decoded = base64.b64decode(file_data)
    xls = pd.ExcelFile(io.BytesIO(decoded))

    selected_sheet_name = xls.sheet_names[idx]   # ⭐ STRING VARIABLE
    df = pd.read_excel(xls, sheet_name=selected_sheet_name)
    df.columns = df.columns.str.strip()

    company_col = "Company" if "Company" in df.columns else df.columns[0]
    options = [{"label": "All Clients", "value": "ALL_COMPANIES"}] + [
        {"label": c, "value": c} for c in sorted(df[company_col].dropna().unique())
    ]

    classes = [
        "sheet-btn selected" if i == idx else "sheet-btn"
        for i in range(len(n_clicks))
    ]

    controls = html.Div(className="glass", children=[
        html.H3("Filter & Analysis Controls"),

        html.Div(className="filter-grid", children=[
            dcc.Dropdown(
                id="company-dropdown",
                options=options,
                value="ALL_COMPANIES",
                className="modern-dropdown"
            ),
            dcc.Input(
                id="start-date-input",
                placeholder="Start Date (DD/MM/YYYY)",
                className="input-box"
            ),
            dcc.Input(
                id="end-date-input",
                placeholder="End Date (DD/MM/YYYY)",
                className="input-box"
            )
        ]),

        html.Div(className="button-row", children=[
            html.Button("Compute Analysis", id="compute-button", className="btn-primary"),
            html.Button("Download PPT", id="download-ppt-button", className="btn-secondary")
        ]),

        html.Div(
            f"Selected Sheet: {selected_sheet_name}",
            className="selected-sheet-badge"
        )
    ])

    return (
        controls,
        df.to_json(date_format="iso", orient="split"),
        selected_sheet_name,   # ⭐ STORED AS STRING
        classes
    )

# ============================================================
# toggle to fix the css div problem link with layout and the compute button 
# ============================================================
@app.callback(
    Output("charts-wrapper", "style"),
    Input("compute-button", "n_clicks"),
    prevent_initial_call=True
)
def show_charts(n_clicks):
    if n_clicks:
        return {"display": "block"}
    return {"display": "none"}


# =========================================================================================================================================================
#                                                                   Charts + Call back modules 
# =========================================================================================================================================================

# =============================================================================================
# - Total tickets
# - User-generated incidents
# - Unassigned incidents

@app.callback(
    Output("total-tickets-summary", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def show_total_tickets(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet
):
    if not n_clicks or df_json is None:
        return ""

    df = pd.read_json(df_json, orient="split")

    if "Number" not in df.columns or "Opened" not in df.columns:
        return ""

    # Deduplicate
    df = df.drop_duplicates(subset="Number")

    # Date filtering
    df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")
    if start_date:
        df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
    if end_date:
        df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

    # Company filter
    if selected_company and selected_company != "ALL_COMPANIES" and "Company" in df.columns:
        df = df[df["Company"] == selected_company]

    total_tickets = df["Number"].nunique()

    # ✅ Heading text logic
    sheet_name = (selected_sheet or "Incident").strip()
    heading = f"Total {sheet_name}"

    return html.Div(
        [
            html.H4(
                heading,
                style={
                    "textAlign": "center",
                    "color": SECONDARY_COLOR,
                    "fontSize": "1.6rem",
                    "fontWeight": "700",
                    "marginBottom": "8px"
                }
            ),
            html.Div(
                f"{total_tickets}",
                style={
                    "textAlign": "center",
                    "color": SECONDARY_COLOR,
                    "fontSize": "2.4rem",
                    "fontWeight": "800"
                }
            )
        ]
    )

@app.callback(
    Output("usergen-incident-summary", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def show_usergen_incidents(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet
):
    if not n_clicks or df_json is None:
        return ""

    df = pd.read_json(df_json, orient="split")

    required_cols = {"Number", "Opened", "Opened By"}
    if not required_cols.issubset(df.columns):
        return ""

    df = df.drop_duplicates(subset="Number")

    df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")
    if start_date:
        df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
    if end_date:
        df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

    if selected_company and selected_company != "ALL_COMPANIES" and "Company" in df.columns:
        df = df[df["Company"] == selected_company]

    user_gen_count = df[
        df["Opened By"].notna() &
        (df["Opened By"].astype(str).str.strip() != "")
    ]["Number"].nunique()

    heading = f"User Gen – {(selected_sheet or 'Incident').strip()}"

    return html.Div(
        [
            html.H4(
                heading,
                style={
                    "textAlign": "center",
                    "color": "yellow",   # 🔴 HEADING YELLOW
                    "fontSize": "1.6rem",
                    "fontWeight": "700",
                    "marginBottom": "8px"
                }
            ),
            html.Div(
                f"{user_gen_count}",
                style={
                    "textAlign": "center",
                    "color": "yellow",   # 🔴 VALUE YELLOW
                    "fontSize": "2.4rem",
                    "fontWeight": "800"
                }
            )
        ]
    )

@app.callback(
    Output("unassigned-incident-summary", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def show_unassigned_incidents(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet
):
    if not n_clicks or df_json is None:
        return ""

    df = pd.read_json(df_json, orient="split")

    if not {"Number", "Opened", "Assigned to"}.issubset(df.columns):
        return ""

    # Deduplicate
    df = df.drop_duplicates(subset="Number")

    # Date filtering
    df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")
    if start_date:
        df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
    if end_date:
        df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

    # Company filter
    if selected_company and selected_company != "ALL_COMPANIES" and "Company" in df.columns:
        df = df[df["Company"] == selected_company]

    # Unassigned logic
    unassigned_count = df[
        df["Assigned to"].isna() |
        (df["Assigned to"].astype(str).str.strip() == "")
    ]["Number"].nunique()

    heading = f"Auto Resolved – {(selected_sheet or 'Incident').strip()}"

    return html.Div(
        [
            html.H4(
                heading,
                style={
                    "textAlign": "center",
                    "color": "red",
                    "fontSize": "1.6rem",
                    "fontWeight": "700",
                    "marginBottom": "8px"
                }
            ),
            html.Div(
                f"{unassigned_count}",
                style={
                    "textAlign": "center",
                    "color": "red",
                    "fontSize": "2.4rem",
                    "fontWeight": "800"
                }
            )
        ]
    )


@app.callback(
    Output("assigned-unassigned-gauge", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    prevent_initial_call=True
)
def sla_meet_vs_missed(
    n_clicks, df_json, start_date, end_date, selected_company
):
    PRIMARY_COLOR = "#003135"
    SECONDARY_COLOR = "#00e28b"
    RED_COLOR = "#ff4d4d"

    if not n_clicks or df_json is None:
        return html.Div()

    try:
        # =============================
        # LOAD DATA
        # =============================
        df = pd.read_json(df_json, orient="split")

        required_cols = {"Number", "Opened", "Made SLA"}
        if not required_cols.issubset(df.columns):
            return html.Div("Required columns missing", style={"color": "red"})

        df = df.drop_duplicates(subset="Number")
        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")

        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        if selected_company and selected_company != "ALL_COMPANIES" and "Company" in df.columns:
            df = df[df["Company"] == selected_company]

        if df.empty:
            return html.Div()

        # =============================
        # SLA NORMALIZATION (BULLETPROOF)
        # =============================
        def normalize_sla(val):
            if pd.isna(val) or str(val).strip() == "":
                return None   # truly empty → skip

            v = str(val).strip().lower()

            # SLA MET patterns
            if any(x in v for x in ["true", "yes", "met", "sla met", "1", "y"]):
                return True

            # EVERYTHING ELSE = MISSED
            return False

        df["SLA_FINAL"] = df["Made SLA"].apply(normalize_sla)

        # remove only truly empty SLA rows
        df = df[df["SLA_FINAL"].notna()]

        if df.empty:
            return html.Div(
                "SLA column is completely empty after filters",
                style={"color": "orange", "textAlign": "center"}
            )

        # =============================
        # SLA COUNTS
        # =============================
        sla_met_df = df[df["SLA_FINAL"] == True]
        sla_missed_df = df[df["SLA_FINAL"] == False]

        total = df["Number"].nunique()
        met_count = sla_met_df["Number"].nunique()
        missed_count = sla_missed_df["Number"].nunique()

        met_pct = round((met_count / total) * 100, 1) if total else 0
        missed_pct = round((missed_count / total) * 100, 1) if total else 0

        # =============================
        # GAUGES
        # =============================
        fig = go.Figure()

        fig.add_trace(go.Indicator(
            mode="gauge+number",
            value=met_pct,
            domain={"x": [0, 0.48], "y": [0, 1]},
            title={"text": "SLA Met"},
            number={"suffix": "%"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": SECONDARY_COLOR}
            }
        ))

        fig.add_trace(go.Indicator(
            mode="gauge+number",
            value=missed_pct,
            domain={"x": [0.52, 1], "y": [0, 1]},
            title={"text": "SLA Missed"},
            number={"suffix": "%"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": RED_COLOR}
            }
        ))

        fig.update_layout(
            height=260,
            margin=dict(l=10, r=10, t=40, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="white"
        )

        # =============================
        # UI
        # =============================
        return html.Div([
            html.H4(
                "SLA Compliance Overview",
                style={
                    "textAlign": "center",
                    "color": SECONDARY_COLOR,
                    "fontWeight": "700"
                }
            ),
            html.P(f"SLA Met: {met_count}/{total} | SLA Missed: {missed_count}/{total}"),
            dcc.Graph(figure=fig, config={"displayModeBar": False})
        ])

    except Exception as e:
        return html.Div(str(e), style={"color": "red", "textAlign": "center"})

###  below tow are future options

@app.callback(
    Output("total-ci-summary", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def show_total_ci(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet
):
    if not n_clicks or df_json is None:
        return ""

    df = pd.read_json(df_json, orient="split")

    if "No_of_Instances" not in df.columns or "COMPANY Name" not in df.columns:
        return ""

    df["No_of_Instances"] = pd.to_numeric(
        df["No_of_Instances"], errors="coerce"
    ).fillna(0)

    # Date filtering (same as Total Tickets)
    if "Opened" in df.columns:
        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")
        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

    # Company → COMPANY Name mapping
    if selected_company and selected_company != "ALL_COMPANIES":
        df = df[df["COMPANY Name"] == selected_company]

    total_ci = int(df["No_of_Instances"].sum())

    sheet_name = (selected_sheet or "Incident").strip()
    heading = f"Total CI"

    return html.Div(
        [
            html.H4(
                heading,
                style={
                    "textAlign": "center",
                    "color": SECONDARY_COLOR,
                    "fontSize": "1.6rem",
                    "fontWeight": "700",
                    "marginBottom": "8px"
                }
            ),
            html.Div(
                f"{total_ci}",
                style={
                    "textAlign": "center",
                    "color": SECONDARY_COLOR,
                    "fontSize": "2.4rem",
                    "fontWeight": "800"
                }
            )
        ],
        style={"width": "100%"}
    )

@app.callback(
    Output("assigned-unassigned-gauge", "figure"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    prevent_initial_call=True
)
def assigned_vs_unassigned_gauge(n_clicks, df_json, start_date, end_date, selected_company):
    if not n_clicks or df_json is None:
        return go.Figure()

    df = pd.read_json(df_json, orient="split")

    if not {"Number", "Opened", "Assigned to"}.issubset(df.columns):
        return go.Figure()

    df = df.drop_duplicates(subset="Number")
    df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")

    if start_date:
        df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
    if end_date:
        df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]
    if selected_company and selected_company != "ALL_COMPANIES":
        df = df[df["Company"] == selected_company]

    total = df["Number"].nunique()
    unassigned = df[
        df["Assigned to"].isna() |
        (df["Assigned to"].astype(str).str.strip() == "")
    ]["Number"].nunique()

    unassigned_pct = (unassigned / total * 100) if total else 0

    gauge_color = "red" if unassigned_pct > 50 else SECONDARY_COLOR

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=unassigned_pct,
            number={"suffix": "%", "font": {"size": 26}},
            title={"text": "Unassigned %", "font": {"size": 14}},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": gauge_color},
                "bgcolor": "#0b2f33",
                "borderwidth": 0,
                "steps": [
                    {"range": [0, 100], "color": "#123c40"}
                ],
            }
        )
    )

    fig.update_layout(
        height=220,
        margin=dict(l=10, r=10, t=40, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        font_color="white",
        transition={"duration": 600, "easing": "cubic-in-out"}  # 🎯 animation
    )

    return fig

# =============================================================================================


#============================================================================================================== all company ticket count
@app.callback(
    Output("company-bar-chart", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),   # stored sheet name (string)
    prevent_initial_call=True
)
def generate_company_volume_chart(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):
    if not n_clicks or df_json is None:
        return ""

    try:
        # =============================
        # Load Data
        # =============================
        df = pd.read_json(df_json, orient="split").reset_index(drop=True)

        # =============================
        # Deduplicate
        # =============================
        if "Number" not in df.columns:
            return html.Div("❌ 'Number' column not found.", style={"color": "red"})

        df = df.drop_duplicates(subset="Number", keep="first")

        # =============================
        # Date filtering
        # =============================
        if "Opened" not in df.columns:
            return html.Div("❌ 'Opened' column not found.", style={"color": "red"})

        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")

        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]

        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        # =============================
        # Company filter
        # =============================
        company_col = "Company" if "Company" in df.columns else df.columns[0]

        if selected_company and selected_company != "ALL_COMPANIES":
            df = df[df[company_col] == selected_company]

        if df.empty:
            return html.Div(
                "⚠️ No incidents found for selected filters.",
                style={"color": "orange", "textAlign": "center"}
            )

        # =============================
        # Aggregation (SAFE)
        # =============================
        summary = (
            df[company_col]
            .value_counts()
            .rename_axis("Client")
            .reset_index(name="Count")
        )

        total_tickets = summary["Count"].sum()

        # =============================
        # Dynamic title
        # =============================
        sheet_name = selected_sheet_name or "Incident"
        title_text = f"{sheet_name} Volume by Client"
        subtitle_text = f"Total Tickets: {total_tickets}"

        # =============================
        # Build figure
        # =============================
        fig = px.bar(
            summary,
            x="Client",
            y="Count",
            text="Count",
            color="Client",
            color_discrete_sequence=px.colors.qualitative.Set1
        )

        # --- Bar text styling (YOUR REQUEST) ---
        fig.update_traces(
            textposition="outside",
            textfont=dict(
                size=22,
                family="Arial Black"   # Bold font
            ),
            cliponaxis=False
        )

        # --- Axis labels ---
        fig.update_xaxes(
            title_text="Client",
            tickfont=dict(size=16)
        )

        fig.update_yaxes(
            title_text="Count",
            tickfont=dict(size=16),
            automargin=True
        )

        # --- Layout ---
        fig.update_layout(
            autosize=True,
            margin=dict(l=40, r=40, t=20, b=100),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="white",
            showlegend=False
        )

        # =============================
        # Return neutral chart block
        # =============================
        return html.Div(children=[
            html.H4(
                title_text,
                style={ "textAlign": "center", "color": SECONDARY_COLOR, "fontSize": "1.8rem", "fontWeight": "700", "borderBottom": f"2px solid {SECONDARY_COLOR}", "paddingBottom": "8px" }
            ),
            html.P(
                subtitle_text,
                style={
                    "textAlign": "center",
                    "fontSize": "0.95rem",
                    "opacity": "0.8",
                    "marginBottom": "16px"
                }
            ),
            dcc.Graph(
                figure=fig,
                config={
                    "displayModeBar": False,
                    "responsive": True
                }
            )
        ])

    except Exception as e:
        return html.Div(
            f"❌ Error: {str(e)}",
            style={"color": "red", "textAlign": "center"}
        )

#============================================================================================================== ticket per instance
@app.callback(
    Output("tickets-per-instance-table", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def build_tickets_per_instance_table(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):
    if not n_clicks or df_json is None:
        return ""

    try:
        # =============================
        # Load & normalize data
        # =============================
        df = pd.read_json(df_json, orient="split")
        df.columns = df.columns.str.strip().str.lower()

        required_cols = ["number", "company", "opened"]
        for col in required_cols:
            if col not in df.columns:
                return html.Div(
                    f"Missing required column: {col}",
                    style={"color": "red", "textAlign": "center"}
                )

        # =============================
        # Deduplication & date filter
        # =============================
        df = df.drop_duplicates(subset="number")
        df["opened"] = pd.to_datetime(df["opened"], errors="coerce")

        if start_date:
            df = df[df["opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
        if end_date:
            df = df[df["opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        df["company"] = df["company"].fillna("(blank)").replace("", "(blank)")

        if df.empty:
            return html.Div(
                "No data available for selected filters.",
                style={"color": "orange", "textAlign": "center"}
            )

        # =============================
        # Ticket count per company
        # =============================
        tickets_df = (
            df.groupby("company")["number"]
            .nunique()
            .reset_index()
            .rename(columns={
                "company": "company name",
                "number": "total tickets"
            })
        )

        tickets_df["company name"] = tickets_df["company name"].str.strip()

        # =============================
        # Instance count
        # =============================
        if "company name" in df.columns and "no_of_instances" in df.columns:
            instance_df = (
                df[["company name", "no_of_instances"]]
                .dropna()
                .drop_duplicates(subset=["company name"])
            )
            instance_df["no_of_instances"] = (
                pd.to_numeric(instance_df["no_of_instances"], errors="coerce")
                .fillna(1)
            )
        else:
            instance_df = pd.DataFrame({
                "company name": tickets_df["company name"],
                "no_of_instances": 1
            })

        # =============================
        # Merge & calculate TPI
        # =============================
        summary = pd.merge(
            tickets_df,
            instance_df,
            on="company name",
            how="left"
        )

        summary["no_of_instances"] = summary["no_of_instances"].fillna(1)
        summary["tickets per instance"] = (
            summary["total tickets"] / summary["no_of_instances"]
        ).round(2)

        summary = summary.sort_values(
            ["tickets per instance", "total tickets"],
            ascending=[False, False]
        )

        # =============================
        # Heading
        # =============================
        sheet_name = selected_sheet_name or "Incident"
        heading = f"{sheet_name} – Per Instance Summary"

        # =============================
        # Layout
        # =============================
        return html.Div([

            html.H4(
                heading,
                style={
                    "textAlign": "center",
                    "color": "#00e28b",
                    "fontWeight": "bold",
                    "fontSize": "1.6rem",
                    "borderBottom": "2px solid #00e28b",
                    "paddingBottom": "8px",
                    "marginBottom": "14px",
                }
            ),

            # Download button
            html.Div(
                [
                    html.Button(
                        "Download",
                        id="download-tpi-btn",
                        n_clicks=0,
                        style={
                            "backgroundColor": "#34495e",
                            "color": "white",
                            "border": "none",
                            "padding": "6px 14px",
                            "borderRadius": "6px",
                            "cursor": "pointer",
                            "fontSize": "0.85rem",
                        },
                    ),
                    dcc.Download(id="download-tpi-excel"),
                ],
                style={
                    "display": "flex",
                    "justifyContent": "flex-end",
                    "marginBottom": "10px",
                },
            ),

            dash_table.DataTable(
                id="tpi-datatable",
                data=summary.to_dict("records"),
                columns=[
                    {"name": "Company Name", "id": "company name"},
                    {"name": "No of Instances", "id": "no_of_instances", "type": "numeric"},
                    {"name": "Total Tickets", "id": "total tickets", "type": "numeric"},
                    {"name": "Tickets per Instance", "id": "tickets per instance", "type": "numeric"},
                ],
                style_table={
                    "width": "100%",
                    "overflowX": "auto"
                },
                style_cell={
                    "textAlign": "center",
                    "padding": "10px",
                    "fontSize": "0.95rem",
                    "fontFamily": "Inter, Arial",
                    "color": "#2c3e50",
                },
                style_header={
                    "backgroundColor": "#1abc9c",
                    "color": "white",
                    "fontWeight": "bold",
                    "fontSize": "1rem",
                },
                style_data_conditional=[
                    # Company Name
                    {
                        "if": {"column_id": "company name"},
                        "backgroundColor": "#eef5ff",
                        "color": "#1f3c88",
                        "fontWeight": "600"
                    },
                    # No of Instances
                    {
                        "if": {"column_id": "no_of_instances"},
                        "backgroundColor": "#f4f0ff",
                        "color": "#4b2aad",
                        "fontWeight": "600"
                    },
                    # Total Tickets
                    {
                        "if": {"column_id": "total tickets"},
                        "backgroundColor": "#fff6e5",
                        "color": "#8a5a00",
                        "fontWeight": "600"
                    },
                    # TPI > 1
                    {
                        "if": {
                            "filter_query": "{tickets per instance} > 1",
                            "column_id": "tickets per instance"
                        },
                        "backgroundColor": "#f4b8ad",
                        "color": "red",
                        "fontWeight": "bold",
                    },
                    # TPI <= 1
                    {
                        "if": {
                            "filter_query": "{tickets per instance} <= 1",
                            "column_id": "tickets per instance"
                        },
                        "backgroundColor": "#7fedaf",
                        "color": "green",
                        "fontWeight": "bold",
                    },
                ],
                page_action="none",
                style_as_list_view=True
            )
        ])

    except Exception as e:
        return html.Div(
            f"Error: {str(e)}",
            style={"color": "red", "textAlign": "center"}
        )

@app.callback(
    Output("download-tpi-excel", "data"),
    Input("download-tpi-btn", "n_clicks"),
    State("tpi-datatable", "data"),
    prevent_initial_call=True
)
def download_tpi_excel(n_clicks, table_data):
    if not n_clicks or not table_data:
        return dash.no_update

    df = pd.DataFrame(table_data)

    return dcc.send_data_frame(
        df.to_excel,
        "tickets_per_instance_summary.xlsx",
        index=False
    )

#============================================================================================================== Monthly ticket trend
@app.callback(
    Output("month-bar-chart", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def generate_month_volume_chart(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):
    if not n_clicks or df_json is None:
        return ""

    try:
        # =============================
        # Load & validate data
        # =============================
        df = pd.read_json(df_json, orient="split").reset_index(drop=True)

        if "Number" not in df.columns:
            return html.Div("❌ 'Number' column not found.", style={"color": "red"})

        if "Opened" not in df.columns:
            return html.Div("❌ 'Opened' column not found.", style={"color": "red"})

        df = df.drop_duplicates(subset="Number", keep="first")

        # =============================
        # Date handling
        # =============================
        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")
        df = df.dropna(subset=["Opened"])

        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]

        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        # =============================
        # Company filter
        # =============================
        company_col = "Company" if "Company" in df.columns else None

        if company_col and selected_company and selected_company != "ALL_COMPANIES":
            df = df[df[company_col] == selected_company]

        if df.empty:
            return html.Div(
                "⚠️ No incidents found for selected filters.",
                style={"color": "orange", "textAlign": "center"}
            )

        # =============================
        # Month aggregation (CORRECT)
        # =============================
        df["MonthDate"] = df["Opened"].dt.to_period("M").dt.to_timestamp()

        summary = (
            df.groupby("MonthDate")
              .size()
              .reset_index(name="Count")
              .sort_values("MonthDate")
        )

        summary["Month"] = summary["MonthDate"].dt.strftime("%b %Y")

        total_tickets = summary["Count"].sum()

        # =============================
        # Dynamic heading
        # =============================
        sheet_name = selected_sheet_name or "Incident"
        heading = f"{sheet_name} Monthly Ticket Trend"
        subheading = f"Total Tickets: {total_tickets}"

        # =============================
        # Build figure
        # =============================
        fig = px.bar(
            summary,
            x="Month",
            y="Count",
            text="Count",
            color="Month",
            color_discrete_sequence=px.colors.qualitative.Set1
        )

        fig.update_traces(
            textposition="outside",
            textfont=dict(size=22, family="Arial Black"),
            cliponaxis=False
        )

        fig.update_xaxes(
            categoryorder="array",
            categoryarray=summary["Month"],
            tickangle=-45,
            tickfont=dict(size=16),
            title_text="Month"
        )

        fig.update_yaxes(
            title_text="Count",
            tickfont=dict(size=16),
            automargin=True
        )

        fig.update_layout(
            autosize=True,
            margin=dict(l=40, r=40, t=20, b=100),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="white",
            showlegend=False
        )

        # =============================
        # Return neutral block
        # =============================
        return html.Div(children=[

            html.H4(
                heading,
                style={ "textAlign": "center", "color": SECONDARY_COLOR, "fontSize": "1.8rem", "fontWeight": "700", "borderBottom": f"2px solid {SECONDARY_COLOR}", "paddingBottom": "8px" }
            ),

            html.P(
                subheading,
                className="chart-subtitle"
            ),

            dcc.Graph(
                figure=fig,
                config={
                    "displayModeBar": False,
                    "responsive": True
                }
            )
        ])

    except Exception as e:
        return html.Div(
            f"❌ Error: {str(e)}",
            style={"color": "red", "textAlign": "center"}
        )


#============================================================================================================== MTTR

@app.callback(
    Output("month-clustered-chart", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def generate_month_clustered_chart(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):

    if not n_clicks or df_json is None:
        return ""

    try:
        # =============================
        # Load & validate data
        # =============================
        df = pd.read_json(df_json, orient="split").reset_index(drop=True)

        required_cols = ["Number", "Opened", "Closed"]
        for col in required_cols:
            if col not in df.columns:
                return html.Div(
                    f"❌ '{col}' column not found.",
                    style={"color": "red"}
                )

        # Remove duplicate tickets
        df = df.drop_duplicates(subset="Number", keep="first")

        # =============================
        # Date handling
        # =============================
        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")
        df["Closed"] = pd.to_datetime(df["Closed"], errors="coerce")

        df = df.dropna(subset=["Opened"])

        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]

        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        # =============================
        # Company filter
        # =============================
        if "Company" in df.columns and selected_company and selected_company != "ALL_COMPANIES":
            df = df[df["Company"] == selected_company]

        if df.empty:
            return html.Div(
                "⚠️ No incidents found for selected filters.",
                style={"color": "orange", "textAlign": "center"}
            )

        # =============================
        # Create month columns
        # =============================
        df["OpenedMonth"] = df["Opened"].dt.to_period("M").dt.to_timestamp()
        df["ClosedMonth"] = df["Closed"].dt.to_period("M").dt.to_timestamp()

        # =============================
        # Resolution per ticket
        # =============================
        df["ResolutionHours"] = (
            (df["Closed"] - df["Opened"])
            .dt.total_seconds()
            .div(3600)
        )

        # Fallback rule (missing or negative → 4 hrs)
        df["ResolutionHours"] = df["ResolutionHours"].mask(
            df["ResolutionHours"].isna() | (df["ResolutionHours"] < 0),
            4
        )

        # =============================
        # Opened count
        # =============================
        opened_summary = (
            df.groupby("OpenedMonth")["Number"]
              .nunique()
              .reset_index(name="Opened")
              .rename(columns={"OpenedMonth": "MonthDate"})
        )

        # =============================
        # Closed count
        # =============================
        closed_summary = (
            df.dropna(subset=["Closed"])
              .groupby("ClosedMonth")["Number"]
              .nunique()
              .reset_index(name="Closed")
              .rename(columns={"ClosedMonth": "MonthDate"})
        )

        # =============================
        # MTTR (based on Opened month)
        # =============================
        mttr_summary = (
            df.groupby("OpenedMonth")
              .agg(
                  TotalResolutionHours=("ResolutionHours", "sum"),
                  OpenedCount=("Number", "nunique")
              )
              .reset_index()
              .rename(columns={"OpenedMonth": "MonthDate"})
        )

        mttr_summary["MTTR"] = (
            mttr_summary["TotalResolutionHours"] /
            mttr_summary["OpenedCount"]
        ).round(2)

        # =============================
        # Merge all summaries
        # =============================
        summary = opened_summary.merge(
            closed_summary,
            on="MonthDate",
            how="outer"
        ).merge(
            mttr_summary[["MonthDate", "MTTR"]],
            on="MonthDate",
            how="outer"
        ).fillna(0).sort_values("MonthDate")

        summary["Month"] = summary["MonthDate"].dt.strftime("%b %Y")

        # =============================
        # Dynamic heading
        # =============================
        sheet_name = selected_sheet_name or "Incident"
        heading = f"Monthly {sheet_name} MTTR Trend Distribution"

        # =============================
        # Build figure
        # =============================
        fig = go.Figure()

        # Opened Bar
        fig.add_trace(go.Bar(
            x=summary["Month"],
            y=summary["Opened"],
            name="Tickets Opened",
            marker_color="#2B0BDF",
            text=summary["Opened"],
            textfont=dict(
                color="#2B0BDF",
                size=14
            )
        ))

        # Closed Bar
        fig.add_trace(go.Bar(
            x=summary["Month"],
            y=summary["Closed"],
            name="Tickets Closed",
            marker_color="#9414BB",
            text=summary["Closed"],
            textposition="outside",   # or "inside"
            textfont=dict(
                color="#9414BB",
                size=14
            )
        ))

        # MTTR Line
        fig.add_trace(go.Scatter(
            x=summary["Month"],
            y=summary["MTTR"],
            name="MTTR (Hours)",
            mode="lines+markers+text",
            text=summary["MTTR"],
            textposition="top center",   # important for proper placement
            yaxis="y2",
            line=dict(color="#009b5f", width=4),
            marker=dict(
                size=9,
                color="#009b5f"
            ),
            textfont=dict(
                color=SECONDARY_COLOR,   # 👈 data label color
                size=12
            )
        ))


        # =============================
        # Styling
        # =============================
        fig.update_traces(
            selector=dict(type="bar"),
            textposition="outside",
            textfont=dict(size=22, family="Arial Black"),
            cliponaxis=False
        )

        fig.update_traces(
            selector=dict(type="scatter"),
            textposition="top center",
            textfont=dict(size=16, family="Arial Black")
        )

        fig.update_xaxes(
            tickangle=-45,
            tickfont=dict(size=16),
            title_text="Month"
        )

        fig.update_yaxes(
            title_text="Ticket Count",
            showgrid=False,
            tickfont=dict(size=16),
            range=[0, None],
            automargin=True
        )

        fig.update_layout(
            barmode="group",
            autosize=True,
            margin=dict(l=40, r=40, t=20, b=100),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="white",
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.3,
                xanchor="center",
                x=0.5
            ),
            yaxis2=dict(
                title="MTTR (Hours)",
                overlaying="y",
                side="right",
                showgrid=False
            )
        )

        # =============================
        # Return
        # =============================
        return html.Div(children=[

            html.H4(
                heading,
                style={
                    "textAlign": "center",
                    "color": SECONDARY_COLOR,
                    "fontSize": "1.8rem",
                    "fontWeight": "700",
                    "borderBottom": f"2px solid {SECONDARY_COLOR}",
                    "paddingBottom": "8px"
                }
            ),

            dcc.Graph(
                figure=fig,
                config={
                    "displayModeBar": False,
                    "responsive": True
                }
            )
        ])

    except Exception as e:
        return html.Div(
            f"❌ Error: {str(e)}",
            style={"color": "red", "textAlign": "center"}
        )

#============================================================================================================== State
@app.callback(
    Output("state-pie-chart", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def generate_state_chart(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):
    if not n_clicks or df_json is None:
        return ""

    try:
        # =============================
        # Load & normalize data
        # =============================
        df = pd.read_json(df_json, orient="split").reset_index(drop=True)

        if "Number" not in df.columns:
            return html.Div("❌ 'Number' column not found.", style={"color": "red"})

        df = df.drop_duplicates(subset="Number", keep="first")

        df.columns = df.columns.str.strip()

        if "Opened" not in df.columns:
            return html.Div("❌ 'Opened' column not found.", style={"color": "red"})

        if "State" not in df.columns:
            return html.Div("❌ 'State' column not found.", style={"color": "red"})

        # =============================
        # Date filtering
        # =============================
        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")

        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]

        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        # =============================
        # Company filter
        # =============================
        company_col = "Company" if "Company" in df.columns else None

        if company_col and selected_company and selected_company != "ALL_COMPANIES":
            df = df[df[company_col] == selected_company]

        if df.empty:
            return html.Div(
                "⚠️ No incidents found for selected filters.",
                style={"color": "orange", "textAlign": "center"}
            )

        # =============================
        # Aggregation
        # =============================
        summary = (
            df["State"]
            .astype(str)
            .value_counts()
            .rename_axis("State")
            .reset_index(name="Count")
        )

        total_tickets = summary["Count"].sum()

        # =============================
        # Dynamic heading
        # =============================
        sheet_name = selected_sheet_name or "Incident"
        heading = f"{sheet_name} State Distribution"
        subheading = f"Total Tickets: {total_tickets}"

        # =============================
        # Build figure
        # =============================
        fig = px.bar(
            summary,
            x="State",
            y="Count",
            text="Count",
            color="State",
            color_discrete_sequence=px.colors.qualitative.Plotly
        )

        fig.update_traces(
            textposition="outside",
            textfont=dict(size=22, family="Arial Black"),
            cliponaxis=False,
            hovertemplate="<b>%{x}</b><br>Count: %{y}<extra></extra>"
        )

        fig.update_xaxes(
            title_text="State",
            tickfont=dict(size=16)
        )

        fig.update_yaxes(
            title_text="Count",
            tickfont=dict(size=16),
            automargin=True
        )

        fig.update_layout(
            autosize=True,
            margin=dict(l=40, r=40, t=20, b=100),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="white",
            showlegend=False
        )

        # =============================
        # Return neutral block
        # =============================
        return html.Div(children=[

            html.H4(
                heading,
                style={ "textAlign": "center", "color": SECONDARY_COLOR, "fontSize": "1.8rem", "fontWeight": "700", "borderBottom": f"2px solid {SECONDARY_COLOR}", "paddingBottom": "8px" }
            ),

            html.P(
                subheading,
                className="chart-subtitle"
            ),

            dcc.Graph(
                figure=fig,
                config={
                    "displayModeBar": False,
                    "responsive": True
                }
            )
        ])

    except Exception as e:
        return html.Div(
            f"❌ Error: {str(e)}",
            style={"color": "red", "textAlign": "center"}
        )

#============================================================================================================== Priority of tickets count
@app.callback(
    Output("priority-bar-chart", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def generate_priority_chart(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):
    if not n_clicks or df_json is None:
        return ""

    try:
        # =============================
        # Load & validate data
        # =============================
        df = pd.read_json(df_json, orient="split").reset_index(drop=True)

        required_cols = {"Number", "Opened", "Priority"}
        if not required_cols.issubset(df.columns):
            return html.Div(
                f"❌ Missing columns: {required_cols - set(df.columns)}",
                style={"color": "red"}
            )

        # Deduplicate incidents
        df = df.drop_duplicates(subset="Number", keep="first")

        # =============================
        # Date filtering
        # =============================
        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")

        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        # =============================
        # Company filtering
        # =============================
        company_col = "Company" if "Company" in df.columns else df.columns[0]
        if selected_company and selected_company != "ALL_COMPANIES":
            df = df[df[company_col] == selected_company]

        if df.empty:
            return html.Div(
                "⚠️ No incidents found for selected filters.",
                style={"color": "orange", "textAlign": "center"}
            )

        # =============================
        # Normalize Priority text
        # =============================
        df["Priority"] = (
            df["Priority"]
            .astype(str)
            .str.strip()
            .str.upper()
            .str.replace("–", "-", regex=False)
        )

        # =============================
        # 🔥 UNIFIED PRIORITY LOGIC (KEEP OTHERS AS-IS)
        # =============================
        def map_priority(val):
            if not val or val == "NAN":
                return "UNKNOWN"

            # Already normalized (P1–Pn)
            if val.startswith("P") and val[1:].isdigit():
                return val

            # Numeric based (1 - Critical, 2 - Medium, etc.)
            if val[0].isdigit():
                return f"P{val[0]}"

            # Anything else → keep original
            return val

        df["Priority_Group"] = df["Priority"].apply(map_priority)

        # =============================
        # Aggregation (dynamic order)
        # =============================
        summary = (
            df["Priority_Group"]
            .value_counts()
            .rename_axis("Priority")
            .reset_index(name="Count")
        )

        total_tickets = summary["Count"].sum()

        # =============================
        # Headings
        # =============================
        sheet_name = selected_sheet_name or "Incident"
        heading = f"{sheet_name} Priority Distribution"
        subheading = f"Total Tickets: {total_tickets}"

        # =============================
        # Build bar chart
        # =============================
        fig = px.bar(
            summary,
            x="Priority",
            y="Count",
            text="Count",
            color="Priority",
            color_discrete_sequence=px.colors.qualitative.Set2
        )

        fig.update_traces(
            textposition="outside",
            textfont=dict(size=22, family="Arial Black"),
            cliponaxis=False
        )

        fig.update_xaxes(title_text="Priority", tickfont=dict(size=16))
        fig.update_yaxes(title_text="Count", tickfont=dict(size=16), automargin=True)

        fig.update_layout(
            autosize=True,
            margin=dict(l=40, r=40, t=20, b=100),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="white",
            showlegend=False
        )

        # =============================
        # P1 Ticket List (same logic)
        # =============================
        p1_numbers = (
            df[df["Priority_Group"] == "P1"]["Number"]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        p1_block = []
        if p1_numbers:
            p1_block = [
                html.Div(className="p1-section", children=[
                    html.H5("P1 Tickets", className="p1-title"),
                    html.Div(
                        [html.Div(t, className="p1-item") for t in p1_numbers],
                        className="p1-grid"
                    )
                ])
            ]

        # =============================
        # Return UI
        # =============================
        return html.Div(children=[
            html.H4(
                heading,
                style={
                    "textAlign": "center",
                    "color": SECONDARY_COLOR,
                    "fontSize": "1.8rem",
                    "fontWeight": "700",
                    "borderBottom": f"2px solid {SECONDARY_COLOR}",
                    "paddingBottom": "8px"
                }
            ),
            html.P(subheading, className="chart-subtitle"),
            dcc.Graph(
                figure=fig,
                config={"displayModeBar": False, "responsive": True}
            ),
            *p1_block
        ])

    except Exception as e:
        return html.Div(
            f"❌ Error: {str(e)}",
            style={"color": "red", "textAlign": "center"}
        )


#============================================================================================================== Configuration Item
@app.callback(
    Output("configitem-bar-chart", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    prevent_initial_call=True
)
def update_configitem_chart(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company
):
    if not n_clicks or df_json is None:
        return ""

    try:
        # =============================
        # Load data
        # =============================
        df = pd.read_json(df_json, orient="split").reset_index(drop=True)

        if "Configuration item" not in df.columns:
            return html.Div("❌ 'Configuration item' column not found.", style={"color": "red"})

        if "Opened" not in df.columns:
            return html.Div("❌ 'Opened' column not found.", style={"color": "red"})

        if "Company" not in df.columns:
            return html.Div("❌ 'Company' column not found.", style={"color": "red"})

        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")

        # =============================
        # Date filters
        # =============================
        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]

        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        # =============================
        # ✅ Company filter (KEY FIX)
        # =============================
        if selected_company and selected_company != "ALL_COMPANIES":
            df = df[df["Company"] == selected_company]

        if df.empty:
            return html.Div(
                "⚠️ No incidents found for selected filters.",
                style={"color": "orange", "textAlign": "center"}
            )

        # =============================
        # Aggregation
        # =============================
        summary = (
            df["Configuration item"]
            .value_counts()
            .head(10)
            .rename_axis("Configuration Item")
            .reset_index(name="Count")
        )

        total_servers = df["Configuration item"].nunique()

        # =============================
        # Headings
        # =============================
        heading = "Top 10 Problematic Configuration Items"
        if selected_company and selected_company != "ALL_COMPANIES":
            heading += f" – {selected_company}"

        subheading = f"Total Servers Involved: {total_servers}"

        # =============================
        # Chart (neutral)
        # =============================
        fig = px.bar(
            summary,
            x="Configuration Item",
            y="Count",
            text="Count",
            color="Configuration Item",
            color_discrete_sequence=px.colors.qualitative.Set2
        )

        fig.update_traces(
            textposition="outside",
            textfont=dict(size=22, family="Arial Black"),
            cliponaxis=False
        )

        fig.update_xaxes(
            title_text="Configuration Item",
            tickangle=-40,
            tickfont=dict(size=16)
        )

        fig.update_yaxes(
            title_text="Incident Count",
            tickfont=dict(size=16),
            automargin=True
        )

        fig.update_layout(
            autosize=True,
            margin=dict(l=40, r=40, t=20, b=120),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="white",
            showlegend=False
        )

        return html.Div([
            html.H4(
                heading,
                style={ "textAlign": "center", "color": SECONDARY_COLOR, "fontSize": "1.8rem", "fontWeight": "700", "borderBottom": f"2px solid {SECONDARY_COLOR}", "paddingBottom": "8px" } 
            ),
            html.P(
                subheading,
                style={
                    "textAlign": "center",
                    "fontSize": "0.95rem",
                    "opacity": "0.8",
                    "marginBottom": "16px"
                }
            ),
            dcc.Graph(
                figure=fig,
                config={"displayModeBar": False, "responsive": True}
            )
        ])

    except Exception as e:
        return html.Div(
            f"❌ Error: {str(e)}",
            style={"color": "red", "textAlign": "center"}
        )

#============================================================================================================== TOP PROBLEMS
# ------------------- HELPERS -----------------------------
def categorize_issue(desc):
    if not isinstance(desc, str):
        return "Other"
    desc_lower = desc.lower()
    for group in groups:
        if any(k.lower() in desc_lower for k in group["keywords"]):
            return group["name"]
    return "Other"

def _is_all_companies(val):
    if val is None:
        return True
    return str(val).strip().upper() in {"ALL", "ALL_COMPANIES", "ALL COMPANIES"}


@app.callback(
    Output("config-item-container", "style"),
    Output("config-item-dropdown", "options"),
    Output("config-item-dropdown", "value"),
    Input("compute-button", "n_clicks"),
    State("company-dropdown", "value"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    prevent_initial_call=True
)
def show_config_dropdown(n_clicks, company, df_json, start_date, end_date):

    if not n_clicks or df_json is None or _is_all_companies(company):
        return {"display": "none"}, [], None

    df = pd.read_json(df_json, orient="split")
    df = df.drop_duplicates(subset="Number", keep="first")

    if "Opened" in df.columns:
        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")
        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

    if "Company" in df.columns:
        df = df[df["Company"] == company]

    if "Configuration item" not in df.columns:
        return {"display": "none"}, [], None

    config_items = sorted(df["Configuration item"].dropna().unique())

    options = (
        [{"label": "All Configuration Items", "value": "ALL_CONFIG"}] +
        [{"label": ci, "value": ci} for ci in config_items]
    )

    return {"display": "block", "width": "260px"}, options, "ALL_CONFIG"


# ------------------- MAIN CALLBACK -----------------------
@app.callback(
    Output("top-problems-chart1", "children"),
    Input("compute-button", "n_clicks"),
    Input("config-item-dropdown", "value"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),
    prevent_initial_call=True
)
def generate_top_problems_chart(
    n_clicks,
    selected_config_item,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet
):

    PRIMARY_COLOR = "#003135"
    SECONDARY_COLOR = "#00e28b"

    if not n_clicks or df_json is None:
        return ""

    try:
        # ---------------- LOAD DATA ----------------
        df = pd.read_json(df_json, orient="split").reset_index(drop=True)

        required_cols = {"Number", "Short description", "Opened"}
        if not required_cols.issubset(df.columns):
            return html.Div("Required columns missing.", style={"color": "red"})

        df = df.drop_duplicates(subset="Number", keep="first")
        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")

        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        if not _is_all_companies(selected_company) and "Company" in df.columns:
            df = df[df["Company"] == selected_company]

        if (
            selected_config_item
            and selected_config_item != "ALL_CONFIG"
            and "Configuration item" in df.columns
        ):
            df = df[df["Configuration item"] == selected_config_item]

        if df.empty:
            return html.Div("No data available.", style={"color": "orange"})

        # ---------------- CLEAN SHORT DESCRIPTION ----------------
        if "Configuration item" in df.columns:
            servers = df["Configuration item"].dropna().unique()
            escaped = [re.escape(s) for s in servers]
            pattern = re.compile(rf"^({'|'.join(escaped)})[\s\-]+", re.IGNORECASE)

            df["Cleaned Desc"] = df["Short description"].apply(
                lambda x: pattern.sub("", x).strip() if isinstance(x, str) else x
            )
        else:
            df["Cleaned Desc"] = df["Short description"]

        # ---------------- CATEGORIZE ----------------
        df["Issue"] = df["Cleaned Desc"].apply(categorize_issue)

        # ---------------- TOP 10 (FIXED HERE) ----------------
        summary = (
            df["Issue"]
            .value_counts()
            .head(10)
            .reset_index()
        )

        summary.columns = ["Issue", "Count"]
        summary["Label"] = list(string.ascii_uppercase[:len(summary)])




        # ---------------- FIGURE ----------------
        fig = px.bar(
            summary,
            x="Label",
            y="Count",
            text="Count",
            color="Label",
            color_discrete_sequence=px.colors.qualitative.Set2
        )

        fig.update_traces(
            textposition="outside",
            textfont=dict(size=20, family="Arial Black"),
            cliponaxis=False
        )

        fig.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color="white"),
            showlegend=False,
            margin=dict(l=30, r=30, t=20, b=30),
            xaxis_title="Category",
            yaxis_title="Count",
            yaxis=dict(gridcolor="rgba(255,255,255,0.12)")
        )

        # ---------------- LEGEND GRID ----------------
        legend_items = [
            html.Div(
                [
                    html.Span(
                        f"{row['Label']} – {row['Issue']}",
                        style={"fontWeight": "600", "color": SECONDARY_COLOR}
                    ),
                    html.Button(
                        "⬇",
                        id={"type": "download-btn-problems", "index": row["Label"]},
                        n_clicks=0,
                        style={
                            "background": "transparent",
                            "border": f"1px solid {SECONDARY_COLOR}",
                            "color": SECONDARY_COLOR,
                            "borderRadius": "6px",
                            "padding": "2px 8px",
                            "cursor": "pointer",
                            "fontSize": "0.75rem"
                        }
                    )
                ],
                style={
                    "display": "flex",
                    "justifyContent": "space-between",
                    "alignItems": "center",
                    "gap": "10px"
                }
            )
            for _, row in summary.iterrows()
        ]

        legend_grid = html.Div(
            legend_items,
            style={
                "display": "grid",
                "gridTemplateColumns": "repeat(auto-fit, minmax(260px, 1fr))",
                "gap": "12px",
                "padding": "16px",
                "marginTop": "16px"
            }
        )

        # ---------------- TITLE ----------------
        sheet = selected_sheet or "Incident"
        title = f"Top Problems by Category — {sheet}"

        return html.Div(
            [
                html.H4(
                    title,
                    style={
                        "textAlign": "center",
                        "color": SECONDARY_COLOR,
                        "fontSize": "1.8rem",
                        "fontWeight": "700",
                        "borderBottom": f"2px solid {SECONDARY_COLOR}",
                        "paddingBottom": "8px"
                    }
                ),
                dcc.Graph(figure=fig, config={"displayModeBar": False}),
                legend_grid
            ]
        )

    except Exception as e:
        return html.Div(f"Error: {e}", style={"color": "red"})


# ------------------- DOWNLOAD CALLBACK -------------------
@app.callback(
    Output("download-problems-xlsx", "data"),
    Input({"type": "download-btn-problems", "index": ALL}, "n_clicks"),
    Input("config-item-dropdown", "value"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    prevent_initial_call=True
)
def download_problem_data(
    n_clicks,
    selected_config_item,
    df_json,
    start_date,
    end_date,
    selected_company
):
    if not n_clicks or sum(c or 0 for c in n_clicks) == 0:
        return dash.no_update

    # ---------------- WHICH LABEL WAS CLICKED ----------------
    ctx = callback_context
    label = json.loads(ctx.triggered[0]["prop_id"].split(".")[0])["index"]

    # ---------------- LOAD DATA ----------------
    df = pd.read_json(df_json, orient="split").reset_index(drop=True)
    df = df.drop_duplicates(subset="Number", keep="first")
    df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")

    # ---------------- SAME FILTERS AS CHART ----------------
    if start_date:
        df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
    if end_date:
        df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

    if selected_company and selected_company != "ALL_COMPANIES" and "Company" in df.columns:
        df = df[df["Company"] == selected_company]

    if (
        selected_config_item
        and selected_config_item != "ALL_CONFIG"
        and "Configuration item" in df.columns
    ):
        df = df[df["Configuration item"] == selected_config_item]

    if df.empty:
        return dash.no_update

    # ---------------- CLEAN + CATEGORIZE (SAME AS CHART) ----------------
    if "Configuration item" in df.columns:
        servers = df["Configuration item"].dropna().unique()
        escaped = [re.escape(s) for s in servers]
        pattern = re.compile(rf"^({'|'.join(escaped)})[\s\-]+", re.IGNORECASE)

        df["Cleaned Desc"] = df["Short description"].apply(
            lambda x: pattern.sub("", x).strip() if isinstance(x, str) else x
        )
    else:
        df["Cleaned Desc"] = df["Short description"]

    df["Issue"] = df["Cleaned Desc"].apply(categorize_issue)

    # ---------------- REBUILD SAME SUMMARY ----------------
    summary = (
        df["Issue"]
        .value_counts()
        .head(10)
        .reset_index()
    )
    summary.columns = ["Issue", "Count"]
    summary["Label"] = list(string.ascii_uppercase[:len(summary)])

    # ---------------- MAP LABEL → ISSUE ----------------
    issue_name = summary.loc[summary["Label"] == label, "Issue"]

    if issue_name.empty:
        return dash.no_update

    issue_name = issue_name.iloc[0]

    df_out = df[df["Issue"] == issue_name]

    if df_out.empty:
        return dash.no_update

    # ---------------- EXPORT ----------------
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Filtered Issues")

    buffer.seek(0)

    filename = (
        f"Top_Problem_{issue_name.replace(' ', '_')}_"
        f"{selected_company if selected_company != 'ALL_COMPANIES' else 'ALL'}.xlsx"
    )

    return dcc.send_bytes(buffer.read(), filename)


#============================================================================================================== peak ticket time of day

@app.callback(
    Output("time-of-day-distribution", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),   # ⭐ selected sheet
    prevent_initial_call=True
)
def update_time_of_day_distribution(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):
    if not n_clicks or df_json is None:
        return ""

    import pandas as pd
    import numpy as np
    import plotly.graph_objects as go
    from datetime import datetime
    from dash import html, dcc

    # =============================
    # Load & validate data
    # =============================
    df = pd.read_json(df_json, orient="split").drop_duplicates().reset_index(drop=True)

    required_cols = {"Opened", "Opened By", "Company"}
    if not required_cols.issubset(df.columns):
        return html.Div("❌ Required columns missing.", style={"color": "red"})

    df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")
    df = df.dropna(subset=["Opened"])

    # =============================
    # Date filters
    # =============================
    if start_date:
        df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
    if end_date:
        df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

    # =============================
    # Company filter
    # =============================
    if selected_company and selected_company != "ALL_COMPANIES":
        df = df[df["Company"] == selected_company]

    if df.empty:
        return html.Div(
            "⚠️ No data available for selected filters.",
            style={"color": "orange", "textAlign": "center"}
        )

    # =============================
    # Ticket / Incident type
    # =============================
    df["generation_type"] = df["Opened By"].apply(
        lambda x: "System Generated" if pd.isna(x) or str(x).strip() == "" else "User Generated"
    )

    # =============================
    # Counts
    # =============================
    system_count = (df["generation_type"] == "System Generated").sum()
    user_count = (df["generation_type"] == "User Generated").sum()
    overall_count = len(df)

    # =============================
    # Time binning (15 min)
    # =============================
    df["hour_decimal"] = df["Opened"].dt.hour + df["Opened"].dt.minute / 60
    bins = np.arange(0, 24, 0.25)
    df["time_bin"] = pd.cut(
        df["hour_decimal"],
        bins=bins,
        right=False,
        labels=bins[:-1]
    ).astype(float)

    def to_hhmm(val):
        h = int(val)
        m = int(round((val - h) * 60))
        return f"{h:02d}:{m:02d}"

    grouped = df.groupby(["time_bin", "generation_type"]).size().reset_index(name="count")
    total = df.groupby("time_bin").size().reset_index(name="count")
    total["generation_type"] = "Overall"

    combined = pd.concat([grouped, total], ignore_index=True)

    full_bins = pd.DataFrame({"time_bin": bins[:-1]})
    combined = full_bins.merge(combined, on="time_bin", how="left")
    combined["count"] = combined["count"].fillna(0)
    combined["hover_time"] = combined["time_bin"].apply(to_hhmm)

    # =============================
    # Chart
    # =============================
    color_map = {
        "System Generated": "#2ECC71",
        "User Generated": "#E74C3C",
        "Overall": "#3498DB"
    }

    fig = go.Figure()

    for t in ["System Generated", "User Generated", "Overall"]:
        subset = combined[combined["generation_type"] == t]
        fig.add_trace(go.Scatter(
            x=subset["time_bin"],
            y=subset["count"],
            mode="lines+markers+text" if t == "Overall" else "lines+markers",
            name=t,
            line=dict(color=color_map[t], width=2),
            marker=dict(size=5),
            text=subset["count"] if t == "Overall" else None,
            textposition="top center",
            textfont=dict(
                color="#00e28b",
                size=16 if t == "Overall" else 10,
                family="Arial Black" if t == "Overall" else "Arial"
            ),
            customdata=subset[["hover_time"]],
            hovertemplate="<b>Time:</b> %{customdata[0]}<br><b>Count:</b> %{y}<extra></extra>"
        ))


    fig.update_layout(
        xaxis=dict(
        title="Hour of Day (0–23)",
        tickmode="array",
        tickvals=np.arange(0, 24, 1),
        
        range=[-0.25, 23.75],
        gridcolor='rgba(200, 200, 200, 0.12)',
        zeroline=False,
        tickfont=dict(color='white')
        ),
        yaxis=dict(
            title="Count",
            rangemode="tozero",
            gridcolor='rgba(200, 200, 200, 0.12)',
            zeroline=False,
            tickfont=dict(color='white')
        ),
        hoverlabel=dict(
            bgcolor="#1e2f31",
            bordercolor="#00e28b",
            font=dict(color="white", size=14),
            namelength=-1
        ),
        hovermode="x unified",
        showlegend=True,
        margin=dict(l=40, r=40, t=20, b=40),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white")
    )

    # =============================
    # Dynamic heading
    # =============================
    sheet_name = selected_sheet_name or "Incident"
    entity = sheet_name.rstrip("s")

    company_label = (
        "All Companies"
        if not selected_company or selected_company == "ALL_COMPANIES"
        else f"– {selected_company}"
    )

    heading = f"Peak {entity} Time {company_label}"

    # =============================
    # Output
    # =============================
    return html.Div([

        html.H4(heading, 
            style={
                "textAlign": "center",
                "color": SECONDARY_COLOR,
                "fontSize": "1.8rem",
                "fontWeight": "700",
                "borderBottom": f"2px solid {SECONDARY_COLOR}",
                "paddingBottom": "8px",
                "marginBottom": "10px"
            }
        ),

        html.Div([
            html.Div(f"🟢 System Generated {entity}s: {system_count}", className="summary-item system"),
            html.Div(f"🔴 User Generated {entity}s: {user_count}", className="summary-item user"),
            html.Div(f"🔵 Overall {entity}s: {overall_count}", className="summary-item overall"),
        ], className="summary-row"),

        dcc.Graph(
            figure=fig,
            config={"displayModeBar": False, "responsive": True}
        )

    ])

#============================================================================================================== assignment group
@app.callback(
    Output('assignment-group-bar-chart', 'children'),
    Input('compute-button', 'n_clicks'),
    State('stored-df', 'data'),
    State('start-date-input', 'value'),
    State('end-date-input', 'value'),
    State('company-dropdown', 'value'),
    prevent_initial_call=True
)
def generate_assignment_group_chart(n_clicks, df_json, start_date, end_date, selected_company):

    if not n_clicks or df_json is None:
        return ""

    df = pd.read_json(df_json, orient='split').reset_index(drop=True)

    # -------------------------
    # Deduplication
    # -------------------------
    if 'Number' not in df.columns:
        return html.Div("❌ 'Number' column not found.", style={'color': 'red'})
    df = df.drop_duplicates(subset='Number', keep='first')

    # -------------------------
    # Required columns
    # -------------------------
    if 'Opened' not in df.columns or 'Assignment group' not in df.columns:
        return html.Div("❌ Required columns missing.", style={'color': 'red'})

    df['Opened'] = pd.to_datetime(df['Opened'], errors='coerce')

    # -------------------------
    # Date filters
    # -------------------------
    if start_date:
        df = df[df['Opened'] >= datetime.strptime(start_date, '%d/%m/%Y')]
    if end_date:
        df = df[df['Opened'] <= datetime.strptime(end_date, '%d/%m/%Y')]

    # -------------------------
    # Company filter
    # -------------------------
    company_col = 'Company' if 'Company' in df.columns else df.columns[0]
    if selected_company and selected_company != 'ALL_COMPANIES':
        df = df[df[company_col] == selected_company]

    if df.empty:
        return html.Div("No incidents found for selected filters.", style={'color': 'orange'})

    # -------------------------
    # Assignment group counts (HIGH → LOW)
    # -------------------------
    ag_counts = (
        df['Assignment group']
        .value_counts()
        .reset_index()
    )
    ag_counts.columns = ['Assignment Group', 'Ticket Count']
    ag_counts = ag_counts.sort_values('Ticket Count', ascending=True)  # required for horizontal

    total_tickets = ag_counts['Ticket Count'].sum()

    # -------------------------
    # Horizontal bar chart (NO COLOR CHANGE)
    # -------------------------
    fig = px.bar(
        ag_counts,
        y='Assignment Group',
        x='Ticket Count',
        text='Ticket Count',
        orientation='h'
    )

    fig.update_traces(
        textposition='outside',
        textfont=dict(size=14, family='Arial Black'),
        cliponaxis=False
    )

    fig.update_layout(
        showlegend=False,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color=WHITE,
        autosize=True,
        margin=dict(l=260, r=40, t=30, b=40),
        xaxis=dict(
            title='Ticket Count',
            range=[0, None],
            automargin=True
        ),
        yaxis=dict(
            title='Assignment Group',
            automargin=True
        )
    )

    return html.Div([
        html.H4(
            "Incident Volume by Assignment Group",
            style={
                'textAlign': 'center',
                'color': SECONDARY_COLOR,
                'fontSize': '1.8rem',
                'fontWeight': 'bold',
                'borderBottom': f'2px solid {SECONDARY_COLOR}',
                'paddingBottom': '8px',
                'marginBottom': '10px'
            }
        ),
        html.P(
            f"Total Tickets: {total_tickets}",
            style={
                'textAlign': 'center',
                'color': WHITE,
                'fontSize': '1rem',
                'marginBottom': '20px'
            }
        ),
        dcc.Graph(
            figure=fig,
            config={
                'displayModeBar': False,
                'displaylogo': False,
                'modeBarButtonsToRemove': ['lasso2d', 'select2d']
            }
        )
    ])

#============================================================================================================== type of tickets count
@app.callback(
    Output("type-bar-chart", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),   # ⭐ sheet name string
    prevent_initial_call=True
)
def generate_type_chart(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):
    if not n_clicks or df_json is None:
        return ""

    try:
        # =============================
        # Load data
        # =============================
        df = pd.read_json(df_json, orient="split").reset_index(drop=True)

        # =============================
        # Deduplicate
        # =============================
        if "Number" not in df.columns:
            return html.Div("❌ 'Number' column not found.", style={"color": "red"})

        df = df.drop_duplicates(subset="Number", keep="first")

        # =============================
        # Required columns
        # =============================
        if "Opened" not in df.columns:
            return html.Div("❌ 'Opened' column not found.", style={"color": "red"})

        if "Type" not in df.columns:
            return html.Div("❌ 'Type' column not found.", style={"color": "red"})

        # =============================
        # Date filtering
        # =============================
        df["Opened"] = pd.to_datetime(df["Opened"], errors="coerce")

        if start_date:
            df = df[df["Opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]

        if end_date:
            df = df[df["Opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

        # =============================
        # Company filter
        # =============================
        company_col = "Company" if "Company" in df.columns else df.columns[0]

        if selected_company and selected_company != "ALL_COMPANIES":
            df = df[df[company_col] == selected_company]

        if df.empty:
            return html.Div(
                "⚠️ No incidents found for selected filters.",
                style={"color": "orange", "textAlign": "center"}
            )

        # =============================
        # Hide chart if all INC
        # =============================
        if df["Number"].astype(str).str.startswith("INC").all():
            return ""

        # =============================
        # Normalize Type column
        # =============================
        df["Type"] = df["Type"].fillna("Blank").replace("", "Blank")

        # =============================
        # Aggregation (SAFE)
        # =============================
        summary = (
            df["Type"]
            .value_counts()
            .rename_axis("Type")
            .reset_index(name="Count")
        )

        total_tickets = summary["Count"].sum()

        # =============================
        # Dynamic heading
        # =============================
        sheet_name = selected_sheet_name or "Incident"
        heading = f"{sheet_name} Type Distribution"
        subheading = f"Total Tickets: {total_tickets}"

        # =============================
        # Build figure
        # =============================
        fig = px.bar(
            summary,
            x="Type",
            y="Count",
            text="Count",
            color="Type",
            color_discrete_sequence=px.colors.qualitative.Set2
        )

        fig.update_traces(
            textposition="outside",
            textfont=dict(
                size=22,
                family="Arial Black"   # Bold font
            ),
            cliponaxis=False
        )

        fig.update_xaxes(
            title_text="Type",
            tickfont=dict(size=16)
        )

        fig.update_yaxes(
            title_text="Count",
            tickfont=dict(size=16),
            automargin=True
        )

        fig.update_layout(
            autosize=True,
            margin=dict(l=40, r=40, t=20, b=100),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font_color="white",
            showlegend=False
        )

        # =============================
        # Return neutral chart block
        # =============================
        return html.Div(children=[

            html.H4(
                heading,
                style={ "textAlign": "center", "color": SECONDARY_COLOR, "fontSize": "1.8rem", "fontWeight": "700", "borderBottom": f"2px solid {SECONDARY_COLOR}", "paddingBottom": "8px" }
            ),

            html.P(
                subheading,
                style={
                    "textAlign": "center",
                    "fontSize": "0.95rem",
                    "opacity": "0.8",
                    "marginBottom": "16px"
                }
            ),

            dcc.Graph(
                figure=fig,
                config={
                    "displayModeBar": False,
                    "responsive": True
                }
            )
        ])

    except Exception as e:
        return html.Div(
            f"❌ Error: {str(e)}",
            style={"color": "red", "textAlign": "center"}
        )

#============================================================================================================== assign to
def normalize_employee_name(name):
    if not isinstance(name, str) or name.strip() == "":
        return "Unassigned"

    base_name = name.lower().split("(")[0].strip()

    name_map = {
        # ======================
        # Susnata
        # ======================
        "Susnata Das": "Susnata Sovalin Das",
        "Susnata DAS": "Susnata Sovalin Das",
        "Susnata Sovalin Das": "Susnata Sovalin Das",
        "Susnata Sovalin": "Susnata Sovalin Das",
        "Susnata Sovalin Das (Unisys)": "Susnata Sovalin Das",

        # ======================
        # Mohammed Tauseef / Tausheef
        # ======================
        "Mohammed Tauseef": "Mohammed Tauseef",
        "Tauseef Mohammed": "Mohammed Tauseef",
        "Mohammed Tausheef": "Mohammed Tauseef",
        "Tausheef Mohammed": "Mohammed Tauseef",
        "Md Tauseef": "Mohammed Tauseef",
        "Md Tausheef": "Mohammed Tauseef",

        # ======================
        # Sumanth
        # ======================
        "Sumanth Hs": "Sumanth HS",
        "Sumanth HS": "Sumanth HS",
        "Sumanth Haluvadi": "Sumanth HS",
        "Sumanth Haluvadi Swamy": "Sumanth HS",
        "Sumanth Haluvadi Swamy (Unisys)": "Sumanth HS",

        # ======================
        # Anoop
        # ======================
        "Anoop Kulkarni": "Anoop S Kulkarni",
        "Anoop S Kulkarni": "Anoop S Kulkarni",
        "Kulkarni Anoop": "Anoop S Kulkarni",
        "Anoop S Kulkarni (Unisys)": "Anoop S Kulkarni",
        
        # ======================
        # Manjunath
        # ======================
        "Manjunath M": "Manjunath M",
        "Manjunath Murugesh": "Manjunath M",
        "Manjunath Murgesh": "Manjunath M",
        "Manjunath M (Unisys)": "Manjunath M",

        # ======================
        # Himaja
        # ======================
        "Himaja Gangasan": "Himaja Gangasani",
        "Himaja Gangasani": "Himaja Gangasani",
        "Himaja Gangasani (Unisys)": "Himaja Gangasani",

        # ======================
        # Vinod Kumar Musti
        # ======================
        "Vinod Kumar Yadav Musti": "Vinod Kumar Musti",
        "VinodKumar Musti": "Vinod Kumar Musti",
        "VinodKumar Musti (Unisys)": "Vinod Kumar Musti",
        "Vinod Kumar Musti": "Vinod Kumar Musti",
        "Vinod Kumar Musti (Unisys)": "Vinod Kumar Musti",

        # ======================
        # Nagaraj
        # ======================
        "Nagaraj Durgappa Naik": "Nagaraj Naik",
        "Nagaraj Naik": "Nagaraj Naik",

        # ======================
        # Amar / Amaranadha
        # ======================
        "Gali Amaranadha": "Amaranadha Gali",
        "Amaranadha Gali": "Amaranadha Gali",

        # ======================
        # Anil
        # ======================
        "Anil Babu Manthrala": "Anil Manthrala",
        "Anil BabuA Manthrala": "Anil Manthrala",
        "Anil Manthrala": "Anil Manthrala",
        "Manthrala Anil Babu": "Anil Manthrala",

        # ======================
        # Sankar
        # ======================
        "Sankar Sahu": "Sankar Prasad Sahu",
        "Sahu Sankar Prasad": "Sankar Prasad Sahu",
        "Sankar Prasad Sahu": "Sankar Prasad Sahu",
        "Sankar Prasad Sahu (Unisys)": "Sankar Prasad Sahu",
    }

    # Try exact and partial matches (first 2–3 words)
    tokens = base_name.split()
    for i in range(len(tokens), 0, -1):
        key = " ".join(tokens[:i])
        if key in name_map:
            return name_map[key]

    # Default: title case fallback
    return " ".join(word.capitalize() for word in tokens)

@app.callback(
    Output("assigned-to-bar-chart", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),   # ⭐ SHEET NAME
    prevent_initial_call=True
)
def update_assigned_to_bar_chart(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):

    if not n_clicks or df_json is None:
        return ""

    # ==============================
    # LOAD & CLEAN DATA
    # ==============================
    df = pd.read_json(df_json, orient="split").reset_index(drop=True)
    df.columns = df.columns.str.strip().str.lower()

    required_cols = ["assigned to", "company", "opened", "number"]
    for col in required_cols:
        if col not in df.columns:
            return html.Div(f"❌ '{col}' column not found.", style={"color": "red"})

    df = df.drop_duplicates(subset="number")
    df["opened"] = pd.to_datetime(df["opened"], errors="coerce")

    if start_date:
        df = df[df["opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
    if end_date:
        df = df[df["opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

    if selected_company and selected_company != "ALL_COMPANIES":
        df = df[df["company"] == selected_company]

    if df.empty:
        return html.Div("No data available for selected filters.", style={"color": "orange"})

    # ==============================
    # NORMALIZATION
    # ==============================
    df["assigned to"] = df["assigned to"].replace("", np.nan).fillna("Unassigned")
    df["assigned to"] = df["assigned to"].astype(str).apply(normalize_employee_name)
    df["company"] = df["company"].replace("", np.nan).fillna("(blank)")

    # ==============================
    # AGGREGATION
    # ==============================
    grouped = df.groupby(["assigned to", "company"]).size().reset_index(name="Count")
    total_tickets = df.groupby("assigned to").size().reset_index(name="Total Tickets")

    grouped = grouped.merge(total_tickets, on="assigned to", how="left")
    grouped = grouped.sort_values(["Total Tickets", "assigned to"], ascending=[False, True])

    category_order = grouped["assigned to"].unique()

    # ==============================
    # CHART
    # ==============================
    fig = px.bar(
        grouped,
        x="assigned to",
        y="Count",
        color="company",
        barmode="stack",
        hover_data={
            "Count": True,
            "company": True,
            "Total Tickets": True
        },
        color_discrete_sequence=px.colors.qualitative.Dark24,
        category_orders={"assigned to": category_order}
    )

    # 🔢 Bigger total labels
    for _, row in total_tickets.iterrows():
        fig.add_annotation(
            x=row["assigned to"],
            y=row["Total Tickets"],
            text=str(row["Total Tickets"]),
            showarrow=False,
            font=dict(color="white", size=15, family="Arial Black"),
            yshift=10
        )

    # ==============================
    # STYLE (MATCHES OTHER MODULES)
    # ==============================
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white"),
        margin=dict(l=40, r=40, t=20, b=110),
        xaxis=dict(
            title="Assigned To",
            tickangle=-45,
            showgrid=False
        ),
        yaxis=dict(
            title="Ticket Count",
            gridcolor="rgba(255,255,255,0.12)",
            rangemode="tozero"
        ),
        legend_title_text="Company"
    )

    # ==============================
    # HEADING (FIXED & DYNAMIC)
    # ==============================
    sheet_name = selected_sheet_name or "Incident"
    heading_text = f"{sheet_name} Volume by Assigned To"

    return html.Div([

        html.Div(
            html.H4(
                heading_text,
                style={
                    "textAlign": "center",
                    "color": SECONDARY_COLOR,
                    "fontSize": "1.8rem",
                    "fontWeight": "bold",
                    "borderBottom": f"2px solid {SECONDARY_COLOR}",
                    "paddingBottom": "8px",
                    "marginBottom": "10px"
                }
            ),
            style={"flexShrink": 0}
        ),

        dcc.Graph(
            figure=fig,
            config={"displayModeBar": False, "responsive": True}
        )
    ])

#============================================================================================================== resolve by
@app.callback(
    Output("resolved-by-bar-chart", "children"),
    Input("compute-button", "n_clicks"),
    State("stored-df", "data"),
    State("start-date-input", "value"),
    State("end-date-input", "value"),
    State("company-dropdown", "value"),
    State("selected-sheet-name", "data"),   # ⭐ Selected Sheet
    prevent_initial_call=True
)
def update_resolved_by_bar_chart(
    n_clicks,
    df_json,
    start_date,
    end_date,
    selected_company,
    selected_sheet_name
):

    if not n_clicks or df_json is None:
        return ""

    # =============================
    # LOAD & PREPARE DATA
    # =============================
    df = pd.read_json(df_json, orient="split").reset_index(drop=True)
    df.columns = df.columns.str.strip().str.lower()

    if "resolved by" not in df.columns:
        return ""

    required_cols = ["number", "opened", "company"]
    for col in required_cols:
        if col not in df.columns:
            return html.Div(f"❌ '{col}' column not found.", style={"color": "red"})

    df = df.drop_duplicates(subset="number")
    df["opened"] = pd.to_datetime(df["opened"], errors="coerce")

    if start_date:
        df = df[df["opened"] >= datetime.strptime(start_date, "%d/%m/%Y")]
    if end_date:
        df = df[df["opened"] <= datetime.strptime(end_date, "%d/%m/%Y")]

    if selected_company and selected_company != "ALL_COMPANIES":
        df = df[df["company"] == selected_company]

    if df.empty:
        return html.Div("No data available for selected filters.", style={"color": "orange"})

    # =============================
    # CLEAN & NORMALIZE
    # =============================
    df["resolved by"] = df["resolved by"].replace("", np.nan)
    df = df.dropna(subset=["resolved by"]).copy()

    df["resolved by"] = df["resolved by"].astype(str).apply(normalize_employee_name)
    df["company"] = df["company"].astype(str)

    # =============================
    # AGGREGATION
    # =============================
    grouped = (
        df.groupby(["resolved by", "company"])
        .size()
        .reset_index(name="Count")
    )

    total_tickets = (
        df.groupby("resolved by")
        .size()
        .reset_index(name="Total Tickets")
    )

    grouped = grouped.merge(total_tickets, on="resolved by", how="left")
    grouped = grouped.sort_values(
        ["Total Tickets", "resolved by"], ascending=[False, True]
    )

    category_order = grouped["resolved by"].unique()

    # =============================
    # CHART
    # =============================
    fig = px.bar(
        grouped,
        x="resolved by",
        y="Count",
        color="company",
        barmode="stack",
        hover_data={
            "Count": True,
            "company": True,
            "Total Tickets": True
        },
        color_discrete_sequence=px.colors.qualitative.Dark24,
        category_orders={"resolved by": category_order}
    )

    # 🔢 Bigger total labels (same as Assigned To)
    for _, row in total_tickets.iterrows():
        fig.add_annotation(
            x=row["resolved by"],
            y=row["Total Tickets"],
            text=str(row["Total Tickets"]),
            showarrow=False,
            font=dict(color="white", size=15, family="Arial Black"),
            yshift=10
        )

    # =============================
    # STYLE (MATCHES ALL MODULES)
    # =============================
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="white"),
        margin=dict(l=40, r=40, t=20, b=110),
        xaxis=dict(
            title="Resolved By",
            tickangle=-45,
            showgrid=False
        ),
        yaxis=dict(
            title="Ticket Count",
            gridcolor="rgba(255,255,255,0.12)",
            rangemode="tozero"
        ),
        legend_title_text="Company"
    )

    # =============================
    # HEADING (FIXED + DYNAMIC)
    # =============================
    sheet_name = selected_sheet_name or "Incident"
    heading_text = f"{sheet_name} Volume by Resolved By"

    return html.Div([

        html.H4(
            heading_text,
            style={
                "textAlign": "center",
                "color": SECONDARY_COLOR,
                "fontSize": "1.8rem",
                "fontWeight": "bold",
                "borderBottom": f"2px solid {SECONDARY_COLOR}",
                "paddingBottom": "8px",
                "marginBottom": "12px"
            }
        ),

        dcc.Graph(
            figure=fig,
            config={"displayModeBar": False, "responsive": True}
        )
    ])



#============================================================================================================== END chart 




# ============================================================
if __name__ == "__main__":
    url = "http://127.0.0.1:8050/"

    Timer(1, lambda: webbrowser.open(url)).start()

    app.run(
        debug=False,
        port=8050,
        use_reloader=False
    )

