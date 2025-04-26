# ADD YOUR CLIENT-ID AND TENANT-ID AROUND LINE 28
# Core Streamlit and Data Handling Libraries
import streamlit as st
import pandas as pd
import time # Import time for potential debugging pauses
from datetime import datetime # <<<--- ENSURED THIS IMPORT IS PRESENT

# Authentication Library
from msal_streamlit_authentication import msal_authentication # Streamlit component for MSAL

# Standard Libraries for API Interaction and Data Processing
import requests
import base64
import gzip
import urllib.parse
import json
import os

# Optional: Load environment variables for local development
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # dotenv not installed, proceed without it
    pass

# --- Configuration ---
# Load Azure AD App Registration details from environment variables or use placeholders
CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "PUT-YOUR-CLIENT-ID-HERE")
TENANT_ID = os.getenv("AZURE_TENANT_ID", "PUT-YOUR-TENANT-ID-HERE")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Scopes required for Microsoft Graph (User Info) and Azure Management (Resource Graph, Log Analytics)
SCOPES = ["https://management.azure.com/user_impersonation", "User.Read"]

# Redirect URI must match one registered in your Azure AD App Registration.
REDIRECT_URI = os.getenv("REDIRECT_URI", "http://localhost:8503") # Port 8503

# Azure API Versions
ARG_API_VERSION = "2021-03-01"
# Use the API version confirmed to work
LOG_ANALYTICS_API_VERSION = "2017-10-01"


# --- Logging Function ---
# Centralized function to add logs to session state
def add_log(message, level="info"):
    """Adds a log message with a timestamp and level to the session state."""
    if 'log_messages' not in st.session_state:
        st.session_state.log_messages = []
    # Now 'datetime' is defined because of the import at the top
    timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3] # HH:MM:SS.ms
    st.session_state.log_messages.append({"timestamp": timestamp, "level": level, "message": message})

# --- Dark Mode CSS (Refined) ---
dark_mode_css = """
<style>
    /* Base */
    body {
        color: #E0E0E0; /* Slightly off-white for less harshness */
        background-color: #121212; /* Standard dark background */
    }
    /* Main container */
    .main .block-container {
        background-color: #121212;
    }
    /* Sidebar */
    .stSidebar {
        background-color: #1E1E1E; /* Slightly lighter dark */
        border-right: 1px solid #333; /* Subtle border */
    }
     .stSidebar [data-testid="stMarkdownContainer"] p,
     .stSidebar [data-testid="stMarkdownContainer"] h1,
     .stSidebar [data-testid="stMarkdownContainer"] h2,
     .stSidebar [data-testid="stMarkdownContainer"] h3,
     .stSidebar [data-testid="stMarkdownContainer"] li,
     .stSidebar label { /* Target labels too */
        color: #E0E0E0;
    }
    /* Titles and Headers */
    h1, h2, h3, h4, h5, h6 {
        color: #FFFFFF; /* Pure white for titles */
    }
    /* Buttons */
    .stButton > button {
        background-color: #007BFF; /* Primary blue */
        color: #FFFFFF;
        border: 1px solid #007BFF;
        border-radius: 5px;
        padding: 0.4rem 0.8rem;
    }
    .stButton > button:hover {
        background-color: #0056b3;
        border-color: #0056b3;
    }
    .stButton > button:disabled {
        background-color: #444;
        border-color: #555;
        color: #888;
    }
    /* Download Button */
    .stDownloadButton > button {
        background-color: #17A2B8; /* Info blue */
        color: #FFFFFF;
        border: 1px solid #17A2B8;
        border-radius: 5px;
        padding: 0.4rem 0.8rem;
    }
     .stDownloadButton > button:hover {
        background-color: #117a8b;
        border-color: #117a8b;
    }
    /* Input Fields */
    .stTextInput input, .stTextArea textarea {
        background-color: #2A2A2A;
        color: #E0E0E0;
        border: 1px solid #444;
        border-radius: 5px;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
         border-color: #007BFF; /* Highlight focus */
         box-shadow: 0 0 0 1px #007BFF;
    }
    /* Selectbox / Dropdowns */
    .stSelectbox div[data-baseweb="select"] > div {
        background-color: #2A2A2A;
        color: #E0E0E0;
        border: 1px solid #444;
        border-radius: 5px;
    }
    .stSelectbox div[data-baseweb="select"] > div:hover {
         border-color: #555;
    }
    /* Dataframe */
    .stDataFrame {
        background-color: #1E1E1E; /* Match sidebar */
        border: 1px solid #333;
        border-radius: 5px;
    }
    .stDataFrame thead th {
        background-color: #2A2A2A; /* Darker header */
        color: #FFFFFF;
        border-bottom: 1px solid #444;
    }
    .stDataFrame tbody td {
        color: #E0E0E0;
        border-color: #333;
    }
    /* Status messages (Alerts, Info, etc.) */
    .stAlert { border-radius: 5px; border-width: 1px; border-style: solid; }
    .stAlert[data-baseweb="alert"][kind="info"] { background-color: #11324D; color: #B0E0E6; border-color: #17A2B8; }
    .stAlert[data-baseweb="alert"][kind="success"] { background-color: #144D2E; color: #C3E6CB; border-color: #28A745; }
    .stAlert[data-baseweb="alert"][kind="warning"] { background-color: #4D4D00; color: #FFF3CD; border-color: #FFC107; }
    .stAlert[data-baseweb="alert"][kind="error"] { background-color: #4D1A1A; color: #F8D7DA; border-color: #DC3545; }
    .stAlert a { color: inherit; font-weight: bold; } /* Make links in alerts stand out */

    /* Radio buttons (Theme Toggle) */
    .stRadio [role="radiogroup"] label {
        background-color: #2A2A2A;
        color: #E0E0E0;
        border: 1px solid #444;
        padding: 5px 10px;
        margin: 2px;
        border-radius: 5px;
        transition: background-color 0.2s ease, border-color 0.2s ease;
    }
     .stRadio [role="radiogroup"] label:hover {
         border-color: #555;
         background-color: #333;
     }
     /* Style the selected radio button */
     .stRadio [role="radiogroup"] input:checked + div {
        color: #FFFFFF; /* Make selected text brighter */
     }

    /* Links */
    a { color: #17A2B8; } /* Use info blue for links */
    a:hover { color: #117a8b; }

    /* Expander for logs */
    .stExpander header {
        background-color: #2A2A2A;
        color: #E0E0E0;
        border-radius: 5px 5px 0 0;
        border: 1px solid #444;
    }
     .stExpander header:hover {
         background-color: #333;
     }
    .stExpander div[data-testid="stExpanderDetails"] {
        background-color: #1E1E1E;
        border: 1px solid #444;
        border-top: none;
        border-radius: 0 0 5px 5px;
        padding: 10px;
    }

</style>
"""

# --- Helper Functions ---

def make_api_request(url, access_token, method="GET", headers=None, json_payload=None, timeout=60):
    """Makes an authenticated request to an Azure API endpoint. Returns response JSON or None."""
    log_prefix = "[make_api_request]"
    add_log(f"Calling {method} {url}", level="debug")
    if headers is None:
        headers = {}
    headers['Authorization'] = f'Bearer {access_token}'
    if method.upper() == "POST" and json_payload is not None:
        headers['Content-Type'] = 'application/json'
        add_log(f"Payload: {json.dumps(json_payload)}", level="debug")

    try:
        if method.upper() == "POST":
            response = requests.post(url, headers=headers, json=json_payload, timeout=timeout)
        else:
            response = requests.get(url, headers=headers, timeout=timeout)

        add_log(f"Response Status Code: {response.status_code}", level="debug")
        response.raise_for_status()
        response_json = response.json() if response.content else {}
        add_log(f"Success. Returning JSON response.", level="info")
        return response_json
    except requests.exceptions.Timeout:
        add_log(f"API Request Timed Out: {url}", level="error")
        return None
    except requests.exceptions.RequestException as e:
        add_log(f"API Request Failed: {e}", level="error")
        if hasattr(e, 'response') and e.response is not None:
            add_log(f"Response Status Code: {e.response.status_code}", level="error")
            try:
                error_details = e.response.json()
                add_log(f"Error Content: {json.dumps(error_details, indent=2)}", level="error")
            except json.JSONDecodeError:
                add_log(f"Error Content (non-JSON): {e.response.text}", level="error")
        return None
    except Exception as e:
        add_log(f"Unexpected error during API request: {e}", level="error")
        return None

# @st.cache_data(ttl=3600) # Keep caching disabled for now
def list_subscriptions(access_token):
    """Lists subscriptions using Azure Resource Graph. Returns list or None."""
    log_prefix = "[list_subscriptions]"
    add_log("Function called.", level="info")
    url = f"https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version={ARG_API_VERSION}"
    query = "ResourceContainers | where type =~ 'microsoft.resources/subscriptions' | project subscriptionId, name=properties.displayName"
    payload = {"query": query, "options": {"resultFormat": "objectArray"}}

    add_log(f"Preparing to call make_api_request for URL: {url}", level="debug")
    response = make_api_request(url, access_token, method="POST", json_payload=payload)
    add_log(f"Received response from make_api_request: Type={type(response)}", level="debug")

    if response and isinstance(response.get('data'), list):
        add_log(f"Successfully listed {len(response['data'])} subscriptions.", level="info")
        return response['data']
    elif response is not None:
        add_log("API call succeeded but no 'data' list found or unexpected format.", level="warning")
        # add_log(f"Response Content: {json.dumps(response)}", level="warning") # Optional: log full response
        return []
    else:
        add_log("API request failed (make_api_request returned None).", level="error")
        return None

# @st.cache_data(ttl=300) # Keep caching disabled for now
def list_sentinel_workspaces(access_token, subscription_id):
    """Lists Sentinel-enabled Log Analytics workspaces using ARG. Returns list or None."""
    log_prefix = "[list_sentinel_workspaces]"
    add_log(f"Function called for subscription: {subscription_id}", level="info")
    url = f"https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version={ARG_API_VERSION}"
    query = """
    Resources
    | where type == 'microsoft.operationsmanagement/solutions' and plan.name contains 'securityinsights'
    | project id = tostring(properties.workspaceResourceId), name = split(properties.workspaceResourceId, '/')[-1],
              location = location, resourceGroup = resourceGroup, subscriptionId = subscriptionId, tenantId = tenantId
    | where isnotempty(id)
    """
    payload = {"subscriptions": [subscription_id], "query": query, "options": {"resultFormat": "objectArray"}}

    add_log(f"Preparing to call make_api_request with Solution query (v4) for URL: {url}", level="debug")
    response = make_api_request(url, access_token, method="POST", json_payload=payload)
    add_log(f"Received response from make_api_request: Type={type(response)}", level="debug")

    workspaces = []
    if response and isinstance(response.get('data'), list):
        for ws in response['data']:
            if all(k in ws and ws[k] for k in ['id', 'name', 'resourceGroup', 'subscriptionId', 'location']):
                workspaces.append({
                    "id": ws['id'], "name": ws['name'], "resourceGroup": ws['resourceGroup'],
                    "subscriptionId": ws['subscriptionId'], "location": ws['location'],
                    "tenantId": ws.get('tenantId', TENANT_ID)
                })
            else:
                add_log(f"Skipping workspace due to missing details projected by query: {ws}", level="warning")
        add_log(f"Successfully listed {len(workspaces)} workspaces.", level="info")
        return workspaces
    elif response is not None:
        add_log("API call succeeded but no 'data' list found or unexpected format.", level="warning")
        # add_log(f"Response Content: {json.dumps(response)}", level="warning")
        return []
    else:
        add_log("API request failed (make_api_request returned None).", level="error")
        return None

def execute_kql_query(access_token, workspace_details, kql_query, timespan="P1D"):
    """Executes a KQL query. Returns response JSON or None."""
    log_prefix = "[execute_kql_query]"
    add_log(f"Function called for workspace {workspace_details.get('name', 'N/A')}", level="info")
    workspace_id = workspace_details['id']
    url = f"https://management.azure.com{workspace_id}/query?api-version={LOG_ANALYTICS_API_VERSION}"
    payload = {"query": kql_query, "timespan": timespan}
    add_log(f"Preparing to call make_api_request for URL: {url} (API Version: {LOG_ANALYTICS_API_VERSION})", level="debug")
    response = make_api_request(url, access_token, method="POST", json_payload=payload, timeout=180)
    add_log(f"Received response from make_api_request: Type={type(response)}", level="debug")
    return response

def encode_query_for_deep_link(query: str) -> str:
    """Encodes a KQL query for Azure Portal deep link."""
    try:
        encoded_query = urllib.parse.quote_plus(base64.b64encode(gzip.compress(query.encode('utf-8'))).decode('utf-8'))
        return encoded_query
    except Exception as e:
        add_log(f"Error encoding query for deep link: {e}", level="error")
        return ""

# --- Streamlit App UI ---

st.set_page_config(page_title="Sentinel KQL Hunter", layout="wide")

# --- State Initialization ---
init_state = {
    'auth_data': None, 'access_token': None, 'subscriptions': None,
    'selected_subscription_id': None, 'workspaces': None, 'selected_workspace_details': None,
    'kql_query': "", 'query_results_df': None, 'deep_link_url': None,
    'last_executed_query': "", 'theme_mode': 'Rookie',
    'log_messages': [] # Initialize log message list
}
for key, val in init_state.items():
    if key not in st.session_state:
        st.session_state[key] = val

# --- Apply Theme ---
if st.session_state.theme_mode == '1337':
    st.markdown(dark_mode_css, unsafe_allow_html=True)

st.title("🛡️ Microsoft Sentinel KQL Query Runner")
add_log(f"Script run starting... Access token: {'Set' if st.session_state.access_token else 'None'}", level="debug")

# --- Authentication ---
if CLIENT_ID == "YOUR_CLIENT_ID_HERE" or TENANT_ID == "YOUR_TENANT_ID_HERE":
    st.error("Azure Client ID and/or Tenant ID are not configured.")
    add_log("Azure Client ID/Tenant ID missing.", level="critical")
    st.stop()

# Check for redirect params before calling msal_authentication
query_params = st.query_params
is_redirect = query_params.get_all("code") or query_params.get_all("error")
if is_redirect: add_log("Detected redirect parameters in URL.", level="debug")

add_log("Calling msal_authentication component...", level="debug")
auth_data = msal_authentication(
    auth={"clientId": CLIENT_ID, "authority": AUTHORITY, "redirectUri": REDIRECT_URI, "postLogoutRedirectUri": REDIRECT_URI},
    cache={"cacheLocation": "sessionStorage", "storeAuthStateInCookie": False},
    login_request={"scopes": SCOPES},
    logout_request={},
    key="msal_auth"
)
add_log(f"msal_authentication returned: Type={type(auth_data)}", level="debug")
# if isinstance(auth_data, dict): # Avoid logging sensitive token details
#     add_log(f"Auth data keys: {list(auth_data.keys())}", level="debug")

# --- Post-Authentication Logic ---
st.session_state.auth_data = auth_data
token_changed = False

if auth_data and 'accessToken' in auth_data:
    # add_log("Access token DETECTED in auth_data.", level="info") # Implicitly logged by sidebar msg
    if st.session_state.access_token != auth_data['accessToken']:
        add_log(f"Token state CHANGED (New token/first login).", level="info")
        st.session_state.access_token = auth_data['accessToken']
        token_changed = True
    # else: add_log("Token state UNCHANGED.", level="debug") # Can be noisy
    account_info = auth_data.get('account', {})
    st.sidebar.success(f"Logged in as: {account_info.get('name', 'N/A')}") # Removed email for brevity
else:
    # add_log("No access token found in auth_data this run.", level="debug") # Can be noisy
    if st.session_state.access_token is not None:
        add_log("Token state CHANGED (Token lost).", level="info")
        st.session_state.access_token = None
        token_changed = True
    # else: add_log("Token state UNCHANGED (already None).", level="debug")

if token_changed:
    add_log("Clearing downstream state due to token change...", level="info")
    st.session_state.subscriptions = None
    st.session_state.selected_subscription_id = None
    st.session_state.workspaces = None
    st.session_state.selected_workspace_details = None
    st.session_state.query_results_df = None
    st.session_state.deep_link_url = None
    add_log("Rerunning script now due to token state change...", level="warning")
    st.rerun()

# --- Main Application Area (Requires Authentication) ---
access_token = st.session_state.access_token

if not access_token:
    st.warning("Please login using the button above (or wait for redirect).")
    add_log("Stopping script: No access token.", level="warning")
    # Display logs collected so far before stopping
    with st.expander("Debug Logs", expanded=False):
        log_container = st.container()
        log_html = ""
        for log in reversed(st.session_state.log_messages): # Show newest first
            color = {"info": "#B0E0E6", "warning": "#FFF3CD", "error": "#F8D7DA", "critical": "#F8D7DA", "debug": "#AAAAAA"}.get(log['level'], "#AAAAAA")
            icon = {"info": "ℹ️", "warning": "⚠️", "error": "❌", "critical": "💥", "debug": "🐞"}.get(log['level'], "➡️")
            # Basic escaping for the message to prevent accidental HTML injection
            message_escaped = log['message'].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            log_html += f'<p style="color:{color}; margin: 0; font-size: 0.9em; white-space: pre-wrap; word-wrap: break-word;"><small>{log["timestamp"]} {icon} {log["level"].upper()}: {message_escaped}</small></p>'
        log_container.markdown(log_html, unsafe_allow_html=True)
    st.stop()

add_log("Proceeding with main logic.", level="info")

# --- Sidebar ---
st.sidebar.header("Azure Configuration")

# 1. Subscription Selection
if st.session_state.subscriptions is None:
    add_log("Subscription state is None, attempting to load...", level="info")
    with st.spinner("Loading Azure subscriptions..."):
        st.session_state.subscriptions = list_subscriptions(access_token)
        if st.session_state.subscriptions is None:
             st.sidebar.error("Failed to load subscriptions.")
             add_log("[UI] list_subscriptions returned None.", level="error")
        elif not st.session_state.subscriptions:
             st.sidebar.warning("No subscriptions found.")
             add_log("[UI] list_subscriptions returned an empty list.", level="warning")
        else:
             st.sidebar.info(f"Loaded {len(st.session_state.subscriptions)} subscriptions.")
             add_log(f"[UI] Loaded {len(st.session_state.subscriptions)} subscriptions.", level="info")
             # add_log(f"[UI] Raw Sub Data: {json.dumps(st.session_state.subscriptions)}", level="debug") # Optional detailed log

sub_options = {"-- Select Subscription --": None}
if isinstance(st.session_state.subscriptions, list) and st.session_state.subscriptions:
    add_log("[UI] Populating sub_options dictionary...", level="debug")
    try:
        for sub in st.session_state.subscriptions:
            sub_id = sub.get('subscriptionId')
            sub_name = sub.get('name')
            if not sub_id:
                add_log(f"Skipping subscription with missing ID: {sub}", level="warning")
                continue
            display_name = sub_name.strip() if sub_name and isinstance(sub_name, str) and sub_name.strip() else f"Subscription ({sub_id})"
            if display_name.startswith("Subscription ("): add_log(f"Sub ID {sub_id} has null/invalid name '{sub_name}'. Using ID.", level="warning")
            sub_options[display_name] = sub_id
        # add_log(f"[UI] sub_options populated: {sub_options}", level="debug")
    except Exception as ex:
        add_log(f"Exception populating sub_options: {ex}", level="error")
        # add_log(f"Problematic Sub Data: {json.dumps(st.session_state.subscriptions)}", level="error")

current_sub_name = next((name for name, sub_id in sub_options.items() if sub_id == st.session_state.selected_subscription_id), "-- Select Subscription --")
if current_sub_name not in sub_options: current_sub_name = "-- Select Subscription --"
options_list = list(sub_options.keys())
add_log(f"[UI] Sub options for selectbox: {options_list}", level="debug")

selected_sub_name = st.sidebar.selectbox("Select Azure Subscription:", options=options_list, index=options_list.index(current_sub_name), key="sub_selectbox")
newly_selected_sub_id = sub_options.get(selected_sub_name)

if newly_selected_sub_id != st.session_state.selected_subscription_id:
    add_log(f"Subscription selection changed to: {selected_sub_name} (ID: {newly_selected_sub_id})", level="info")
    st.session_state.selected_subscription_id = newly_selected_sub_id
    if newly_selected_sub_id is not None or st.session_state.selected_workspace_details is not None:
        st.session_state.workspaces = None; st.session_state.selected_workspace_details = None; st.session_state.query_results_df = None; st.session_state.deep_link_url = None
        st.rerun()

# 2. Workspace Selection
if st.session_state.selected_subscription_id:
    if st.session_state.workspaces is None:
        add_log("Workspace state is None, attempting to load...", level="info")
        sub_name_for_spinner = selected_sub_name if selected_sub_name != "-- Select Subscription --" else st.session_state.selected_subscription_id
        with st.spinner(f"Loading Sentinel workspaces for '{sub_name_for_spinner}'..."):
            st.session_state.workspaces = list_sentinel_workspaces(access_token, st.session_state.selected_subscription_id)
            if st.session_state.workspaces is None: st.sidebar.error("Failed to load workspaces.")
            elif not st.session_state.workspaces: st.sidebar.warning("No Sentinel workspaces found.")
            else: st.sidebar.info(f"Loaded {len(st.session_state.workspaces)} workspaces.")

    ws_options = {"-- Select Workspace --": None}
    if isinstance(st.session_state.workspaces, list) and st.session_state.workspaces:
        ws_options.update({ws.get('name', f"WS_{ws.get('id', 'Unknown')}"): ws for ws in st.session_state.workspaces})

    current_ws_name = "-- Select Workspace --"
    if st.session_state.selected_workspace_details: current_ws_name = st.session_state.selected_workspace_details.get('name', "-- Select Workspace --")
    if current_ws_name not in ws_options: current_ws_name = "-- Select Workspace --"
    ws_options_list = list(ws_options.keys())
    add_log(f"[UI] Workspace options for selectbox: {ws_options_list}", level="debug")

    selected_ws_name = st.sidebar.selectbox("Select Sentinel Workspace:", options=ws_options_list, index=ws_options_list.index(current_ws_name), key="ws_selectbox")
    newly_selected_ws_details = ws_options.get(selected_ws_name)

    if newly_selected_ws_details != st.session_state.selected_workspace_details:
        add_log(f"Workspace selection changed to: {selected_ws_name}", level="info")
        st.session_state.selected_workspace_details = newly_selected_ws_details
        if newly_selected_ws_details is not None or st.session_state.query_results_df is not None:
             st.session_state.query_results_df = None; st.session_state.deep_link_url = None
             st.rerun()

# --- Theme Toggle ---
st.sidebar.markdown("---")
st.sidebar.radio("Select Theme:", options=['Rookie', '1337'], key='theme_mode', horizontal=True)
st.sidebar.markdown("---")

# --- Main Panel: Query Input and Results ---
if st.session_state.selected_workspace_details:
    ws_details = st.session_state.selected_workspace_details
    sub_name_for_caption = selected_sub_name if selected_sub_name != "-- Select Subscription --" else "N/A"
    st.header(f"Querying Workspace: {ws_details['name']}")
    st.caption(f"Subscription: {sub_name_for_caption} | RG: {ws_details['resourceGroup']} | Location: {ws_details['location']}")

    # 3. KQL Input
    st.subheader("Kusto Query Language (KQL) Input")
    st.session_state.kql_query = st.text_area("Enter KQL query:", value=st.session_state.kql_query, height=250, key="kql_input_area", placeholder="Example:\nSecurityEvent | take 10")
    if st.session_state.kql_query:
        st.markdown("---"); st.markdown("**Entered Query:**"); st.code(st.session_state.kql_query, language='sql', line_numbers=True); st.markdown("---")

    # Timespan selection
    timespan_display_map = {"P1D": "Past 24h", "P7D": "Past 7d", "P14D": "Past 14d", "P30D": "Past 30d", "PT1H": "Past 1h", "PT6H": "Past 6h", "PT12H": "Past 12h", "PT24H": "Past 24h (P1D)", "CUSTOM": "Custom (N/I)"}
    selected_timespan_display = st.sidebar.selectbox("Query Timespan:", options=list(timespan_display_map.values()), index=0, key="timespan_select")
    timespan = next(code for code, display in timespan_display_map.items() if display == selected_timespan_display)
    can_execute = timespan != "CUSTOM"
    if not can_execute: st.sidebar.warning("Custom timespan N/I.")

    # 4. Execute Button
    if st.button("🚀 Execute Query", key="execute_button", type="primary", disabled=not can_execute):
        if st.session_state.kql_query and ws_details and access_token and can_execute:
            st.session_state.query_results_df = None; st.session_state.deep_link_url = None
            st.session_state.last_executed_query = st.session_state.kql_query
            add_log(f"Executing query: {st.session_state.kql_query[:100]}...", level="info") # Log start of query
            with st.spinner(f"Executing KQL over {timespan_display_map[timespan]}..."):
                api_response = execute_kql_query(access_token, ws_details, st.session_state.kql_query, timespan)
                # Process response
                if api_response and isinstance(api_response.get('tables'), list) and api_response['tables']:
                    primary_table = api_response['tables'][0]
                    if isinstance(primary_table.get('columns'), list) and isinstance(primary_table.get('rows'), list):
                        columns = [col['name'] for col in primary_table['columns']]
                        df = pd.DataFrame(primary_table['rows'], columns=columns)
                        st.session_state.query_results_df = df
                        st.success(f"Query successful! Found {len(df)} results.")
                        add_log(f"Query successful, {len(df)} results.", level="info")
                    else:
                        st.warning("Unexpected result table format.")
                        add_log("Query executed, unexpected table format.", level="warning")
                        st.session_state.query_results_df = pd.DataFrame()
                elif api_response is not None:
                     st.warning("Unexpected/empty response format.")
                     add_log("Query executed, unexpected/empty response.", level="warning")
                     # add_log(f"Raw Response: {json.dumps(api_response)}", level="debug")
                else: # API error logged by helper
                    st.error("Query execution failed. See logs below.") # Point user to logs

            # Generate deep link
            if st.session_state.last_executed_query:
                encoded_kql = encode_query_for_deep_link(st.session_state.last_executed_query)
                if encoded_kql:
                    tenant_id = ws_details.get('tenantId', TENANT_ID)
                    encoded_resource_path = urllib.parse.quote(ws_details['id'], safe='')
                    st.session_state.deep_link_url = f"https://portal.azure.com#@{tenant_id}/blade/Microsoft_Azure_Monitoring_Logs/LogsBlade/resourceId/{encoded_resource_path}/source/LogsBlade.AnalyticsShareLinkToQuery/q/{encoded_kql}"
                # else: Error logged by encode function

        else: # Warnings for missing prerequisites
             if not can_execute: st.warning("Cannot execute with current timespan.")
             elif not st.session_state.kql_query: st.warning("Please enter a KQL query.")
             elif not ws_details: st.warning("Please select a workspace.")

    # 5. Display Results
    if st.session_state.query_results_df is not None:
        st.subheader("📊 Query Results")
        if not st.session_state.query_results_df.empty:
            st.dataframe(st.session_state.query_results_df, use_container_width=True)
            # 6. Download Button
            try:
                csv_data = st.session_state.query_results_df.to_csv(index=False).encode('utf-8')
                st.download_button("💾 Download CSV", csv_data, "sentinel_query_results.csv", "text/csv", key="download_csv_button")
            except Exception as e:
                 add_log(f"Failed to prepare download: {e}", level="error")
                 st.error(f"Failed to prepare data for download: {e}")
        else: st.info("Query returned 0 results.")
        # 7. Deep Link Button
        if st.session_state.deep_link_url: st.link_button("🔗 Open in Azure Portal", st.session_state.deep_link_url)

elif access_token: # Logged in, but no workspace selected
    st.info("👈 Please select an Azure Subscription and Workspace from the sidebar.")

# --- Footer & Log Expander ---
st.markdown("---")

# Display Logs in an Expander at the bottom
with st.expander("Debug Logs", expanded=False):
    log_container = st.container() # Use a container for better control if needed
    # Simple HTML formatting for logs
    log_html = ""
    # Iterate through logs, potentially reversed to show newest first
    for log in reversed(st.session_state.log_messages):
        color = {"info": "#B0E0E6", "warning": "#FFF3CD", "error": "#F8D7DA", "critical": "#F8D7DA", "debug": "#AAAAAA"}.get(log['level'], "#AAAAAA")
        icon = {"info": "ℹ️", "warning": "⚠️", "error": "❌", "critical": "💥", "debug": "🐞"}.get(log['level'], "➡️")
        # Basic escaping for the message to prevent accidental HTML injection
        message_escaped = log['message'].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        log_html += f'<p style="color:{color}; margin: 0; font-size: 0.9em; white-space: pre-wrap; word-wrap: break-word;"><small>{log["timestamp"]} {icon} {log["level"].upper()}: {message_escaped}</small></p>'

    log_container.markdown(log_html, unsafe_allow_html=True)


st.caption(f"Streamlit Sentinel Runner v1.15 (Theme: {st.session_state.theme_mode}, Query API Version 2017-10-01)")
