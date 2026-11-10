"""Acme Order Care / Pink Edition. Launch: python -m streamlit run app.py"""
import math
import os
import re
from pathlib import Path
from html import escape

import streamlit as st
from docx import Document
from openai import OpenAI

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
DOCS = {
    "Shipping / Delivery": DATA / "Shipping_Delivery_Issues.docx",
    "Refund / Return / Exchange": DATA / "Refund_Return_Exchange.docx",
}
EMBEDDING_MODEL = "text-embedding-ada-002"
DEFAULT_CHAT_MODEL = "gpt-4.1-mini"
DEFAULT_BASE_URL = "https://hkust.azure-api.net/hkust-genai/v1/"

st.set_page_config(page_title="Acme · Order Care", page_icon="📦", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Nunito:wght@600;700;800;900&display=swap');
:root {--ink:#513947;--sub:#8e7381;--rose:#c65d83;--border:#f0dfe7;}
html,body,[class*="css"] {font-family:'DM Sans',sans-serif;}
[data-testid="stAppViewContainer"]{background:radial-gradient(ellipse 620px 360px at 6% 6%,#fce5ed 0%,transparent 75%), radial-gradient(ellipse 510px 350px at 96% 34%,#fdf0f2 0%,transparent 75%),#fffafc;}
.block-container{max-width:1330px;padding-top:1.2rem;padding-bottom:3rem;}
header[data-testid="stHeader"]{background:transparent;}
[data-testid="stSidebar"]{background:#fff4f8;border-right:1px solid #f1dfe7;}
[data-testid="stSidebar"] * {color:#6d4558;}
[data-testid="stSidebar"] input{background:white!important;color:#513947!important;}
[data-testid="stSidebar"] button{border-radius:12px;}
[data-testid="stWidgetLabel"] p{font-weight:650;color:#6d4558;}
[data-testid="stTextInput"] input,[data-testid="stTextArea"] textarea {background:#fff!important;border:1px solid #eddce3!important;border-radius:12px!important;}
[data-testid="stForm"]{background:#fff;border:1px solid #eedee5;border-radius:20px;padding:20px 23px 23px;box-shadow:0 10px 28px #b8718c0d;}
[data-testid="stFormSubmitButton"] button{background:#c9638a!important;border:1px solid #c9638a!important;color:#fff!important;border-radius:11px!important;font-weight:750!important;height:45px;}
[data-testid="stFormSubmitButton"] button:hover{background:#b85179!important;}
[data-testid="stDownloadButton"] button{border:1px solid #e7c6d3!important;background:#fff!important;color:#ae5277!important;border-radius:12px;font-weight:700;}
[data-testid="stTabs"] button{font-weight:700;}
[data-testid="stTabs"] [data-baseweb="tab-highlight"]{background:#c9638a;}
[data-testid="stAlert"]{border-radius:14px;}
.brandnav{display:flex;justify-content:space-between;align-items:center;gap:12px;margin:0 0 18px;}
.brandnav .mark{display:flex;align-items:center;gap:12px;color:#633849;font-weight:800;font-size:16px;letter-spacing:-.35px;}
.brandnav .logo{width:37px;height:37px;border-radius:12px;display:grid;place-items:center;background:#f3cede;font-size:21px;}
.brandnav .right{display:flex;align-items:center;gap:10px;font-size:11px;color:#976e83;font-weight:700;}
.nav-pill{padding:7px 11px;background:white;border:1px solid #f0dce5;border-radius:999px;}
.hero{display:flex;align-items:center;justify-content:space-between;gap:18px;background:linear-gradient(112deg,#f8e2e9,#fff4f7 56%,#fff);border:1px solid #f0d8e2;border-radius:24px;padding:19px 25px 17px;box-shadow:0 14px 40px #bf7f9512;overflow:hidden;}
.hero-copy{flex:1;min-width:280px;}
.hero-kicker{font-size:11px;font-weight:800;letter-spacing:2.1px;color:#b25b7d;text-transform:uppercase;margin:0 0 9px;}
.hero h1{font-family:'Nunito',sans-serif;color:#673b51;line-height:1.1;letter-spacing:-1.15px;font-weight:900;font-size:37px;margin:0 0 9px;}
.hero p{font-size:14px;color:#835f70;margin:0;max-width:475px;line-height:1.7;}
.hero-note{margin-top:15px;display:inline-block;border-radius:100px;background:#fff9fb;border:1px solid #edd6e0;padding:8px 12px;color:#97627b;font-size:12px;font-weight:750;}
.microheading{font-family:'Nunito',sans-serif;font-size:16px;font-weight:900;color:#674155;margin:20px 0 8px;}
.legendbar{display:flex;gap:9px;flex-wrap:wrap;margin:13px 0 0;}
.legend{border:1px solid #efdae3;background:#fff;border-radius:12px;padding:9px 12px;display:flex;align-items:center;gap:8px;font-size:12px;color:#805d70;font-weight:700;}
.legend i{font-style:normal;font-size:16px;}
.flowline{display:flex;gap:8px;align-items:center;background:#fff;border:1px solid #efdee5;padding:12px 15px;border-radius:16px;margin:13px 0 20px;}
.flowstep{flex:1;min-width:0;display:flex;align-items:center;gap:9px;}
.stepdot{width:27px;height:27px;flex:0 0 27px;display:grid;place-items:center;border-radius:10px;background:#f9edf1;color:#b67891;font-size:11px;font-weight:800;}
.stepdot.done{background:#d9eee1;color:#51846b;}
.steptxt{font-size:12px;font-weight:750;color:#795a69;white-space:nowrap;}
.steparrow{color:#e6c1d0;font-size:15px;}
.stagefill{height:5px;border-radius:999px;background:#f6e6ed;margin-top:7px;overflow:hidden;}
.stagefill > div{height:5px;background:linear-gradient(90deg,#e69bb7,#c9638a);border-radius:999px;transition:width .2s ease;}
.area-title{display:flex;align-items:center;gap:8px;margin:1px 0 3px;font-family:'Nunito',sans-serif;color:#623d51;font-size:19px;font-weight:900;}
.area-desc{font-size:12px;color:#977786;margin-bottom:15px;}
.explain{border:1px dashed #e9cbd7;border-radius:13px;background:#fff7fa;padding:12px 13px;color:#8c667a;font-size:12px;line-height:1.65;margin:12px 0;}
.result-kicker{font-size:11px;text-transform:uppercase;letter-spacing:1.4px;font-weight:800;color:#9b7187;margin-bottom:8px;}
.case-card{background:#fff;border:1px solid #ecdae2;border-radius:19px;padding:17px 18px;box-shadow:0 10px 24px #a35e7a0b;margin:12px 0;}
.case-head{display:flex;align-items:center;justify-content:space-between;gap:9px;flex-wrap:wrap;margin-bottom:12px;}
.tag{display:inline-block;font-size:12px;font-weight:800;background:#fce9f0;color:#ad5279;border:1px solid #f2cfdc;padding:7px 11px;border-radius:100px;}
.tag.green{background:#ebf7ef;color:#4b8a63;border-color:#d8eddf;}
.case-meta{font-size:11px;color:#9f7a8a;font-weight:650;}
.evidence{background:#fff;border:1px solid #eedee5;border-radius:14px;padding:12px 14px;margin:12px 0;}
.evidence-title{font-size:13px;font-weight:800;color:#684858;}
.evidence-small{font-size:11px;color:#9d7888;margin-top:5px;}
.evidence-score{background:#f8edf1;color:#a45d7a;padding:5px 8px;border-radius:99px;font-size:11px;font-weight:750;}
.footerline{text-align:center;color:#a78392;font-size:11px;margin:23px 0 0;padding:10px;}
@media(max-width:840px){.hero h1{font-size:29px}.hero{padding:18px}.flowstep{gap:4px}.steptxt{font-size:10px}.brandnav .right{display:none}}
</style>
""", unsafe_allow_html=True)

def clean_text(raw: str) -> str:
    # The supplied DOCX files contain some literal RTF escape characters.
    raw = raw.replace("\\'92", "’").replace("\\'96", "–").replace("\\'965", "–5").replace("\\'967", "–7")
    raw = raw.replace("\\{", "{").replace("\\}", "}")
    raw = re.sub(r"\\f\d+|\\fs\d+|\\cf\d+", "", raw)
    raw = raw.strip(" \\")
    return raw


def load_chunks(path: Path) -> list[dict]:
    """Chunk by short document sections; keep FAQ/policy answers together."""
    doc = Document(str(path))
    paragraphs = [clean_text(p.text) for p in doc.paragraphs]
    paragraphs = [p for p in paragraphs if p and p != "}" and p != "---"]
    chunks = []
    section = "General"
    buffer = []
    def flush():
        if buffer:
            chunks.append({"section": section, "text": "\n".join(buffer), "source": path.name})
            buffer.clear()
    for p in paragraphs:
        is_heading = p.startswith("# ") or p.startswith("**Q:") or p.startswith("**Refund Policy") or p.startswith("**Shipping Policy") or p.startswith("**Template:")
        if is_heading:
            flush()
            section = p.strip("# *\\")
        buffer.append(p)
    flush()
    return [c for c in chunks if not (c["text"].startswith("# ") and "\n" not in c["text"])]


def client_for(key: str, base_url: str) -> OpenAI:
    # HKUST endpoint has been used successfully for embeddings in the prior lab.
    # Other providers may not accept this extra api-key header.
    kwargs = {"api_key": key, "base_url": base_url}
    if "hkust.azure-api.net" in base_url:
        kwargs["default_headers"] = {"api-key": key}
    return OpenAI(**kwargs, timeout=45.0, max_retries=1)


def embed_texts(client: OpenAI, texts: list[str]) -> list[list[float]]:
    result = client.embeddings.create(
        model=EMBEDDING_MODEL, input=texts, encoding_format="float"
    )
    return [list(item.embedding) for item in sorted(result.data, key=lambda x: x.index)]


def cosine(a: list[float], b: list[float]) -> float:
    dot = sum(x * y for x, y in zip(a, b))
    na = math.sqrt(sum(x*x for x in a))
    nb = math.sqrt(sum(x*x for x in b))
    return dot / (na * nb) if na and nb else 0.0


def ensure_index(client: OpenAI, kind: str) -> list[dict]:
    """Keep the two vector indexes separate, just like the Dify branches."""
    cache_key = "index_" + kind
    if cache_key not in st.session_state:
        chunks = load_chunks(DOCS[kind])
        vectors = embed_texts(client, [c["text"] for c in chunks])
        st.session_state[cache_key] = [dict(chunk, vector=vector) for chunk, vector in zip(chunks, vectors)]
    return st.session_state[cache_key]


def classify_email(client: OpenAI, model: str, email_body: str) -> str:
    completion = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": (
                "Classify this customer email into exactly one of two labels: "
                "SHIPPING or REFUND. SHIPPING covers order tracking, late delivery, "
                "shipping cost and missing parcels. REFUND covers refunds, returns, "
                "exchanges, wrong item and damaged product. "
                "Return only SHIPPING or REFUND, nothing else."
            )},
            {"role": "user", "content": email_body},
        ],
    )
    label = (completion.choices[0].message.content or "").strip().upper()
    if label.startswith("SHIPPING"):
        return "Shipping / Delivery"
    if label.startswith("REFUND"):
        return "Refund / Return / Exchange"
    raise ValueError(f"Unexpected classifier response: {label!r}")


def retrieve(client: OpenAI, kind: str, email_body: str, top_k: int = 3) -> list[dict]:
    index = ensure_index(client, kind)
    query_vector = embed_texts(client, [email_body])[0]
    matches = [{**c, "score": cosine(query_vector, c["vector"])} for c in index]
    return sorted(matches, key=lambda x: x["score"], reverse=True)[:top_k]


def draft_reply(client: OpenAI, model: str, name: str, email: str, body: str,
                kind: str, matches: list[dict]) -> str:
    evidence = "\n\n".join(f"[{i}] Source: {m['source']} / {m['section']}\n{m['text']}"
                           for i, m in enumerate(matches, 1))
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": (
                "You are the Acme customer support email assistant. Draft a warm, professional, "
                "concise English email based on the provided knowledge-base excerpts. "
                "Start with 'Hi <customer name>,' and end with 'Best,\\nThe Acme Support Team'. "
                "Use only the supplied policy facts; if information is missing, ask for relevant "
                "details (such as order number). Never invent an order status, shipment status, "
                "refund approval, refund processing, or completed action. "
                "Any existing template that makes claims about a specific order is not proof "
                "that the action occurred. Do not promise actions not supported by the policy. "
                "Return only the ready-to-send email, without a subject line or explanation."
            )},
            {"role": "user", "content": (
                f"Category: {kind}\nCustomer name: {name}\nCustomer email: {email}\n"
                f"Customer message:\n{body}\n\nKnowledge base excerpts:\n{evidence}"
            )},
        ],
    )
    return (response.choices[0].message.content or "").strip()




# UI state: sampling uses Streamlit session state instead of hardcoded form defaults.
EXAMPLES = {
    "delivery": {"name": "Megan", "email": "megan@example.com", "body": "My order is delayed, and the tracking information hasn't updated since last week. What should I do?"},
    "wrong_item": {"name": "Marcus", "email": "marcus@example.com", "body": "I ordered a SmartDesk Pro 2, but received a chair instead. How can I return it and receive the correct item?"},
}
for key, value in [("customer_name", EXAMPLES["delivery"]["name"]), ("customer_email", EXAMPLES["delivery"]["email"]), ("customer_body", EXAMPLES["delivery"]["body"])]:
    st.session_state.setdefault(key, value)

with st.sidebar:
    st.markdown("### 📦 Acme Order Care")
    st.caption("Workspace connection")
    api_key = st.text_input("HKUST GenAI Primary Key", value=os.getenv("HKUST_GENAI_API_KEY", ""), type="password")
    with st.expander("Advanced API settings"):
        base_url = st.text_input("API base URL", value=os.getenv("GENAI_BASE_URL", DEFAULT_BASE_URL))
        chat_model = st.text_input("Chat model", value=os.getenv("GENAI_CHAT_MODEL", DEFAULT_CHAT_MODEL))
    st.divider()
    st.markdown("**Knowledge collections**")
    for label, path in DOCS.items():
        st.markdown(("✅ " if path.exists() else "⚠️ ") + label)
    if st.button("Rebuild knowledge index", use_container_width=True, help="Recalculate stored document embeddings for the current session"):
        for key in DOCS:
            st.session_state.pop("index_" + key, None)
        st.success("Index cache cleared. Documents will be indexed on the next request.")
    st.divider()
    st.caption("Keep your API key private, especially when recording your demo.")

st.markdown('''<div class="brandnav"><div class="mark"><span class="logo">📦</span><span>acme<span style="color:#c9638a">.</span> &nbsp; / &nbsp; Order Care</span></div><div class="right"><span class="nav-pill">Customer Operations</span><span class="nav-pill">Support workspace</span></div></div>''', unsafe_allow_html=True)

hero_col, art_col = st.columns([1.55, 0.85], vertical_alignment="center", gap="medium")
with hero_col:
    st.markdown('''<div style="padding:20px 0 8px 5px"><div class="hero-kicker">THE ORDER CARE DESK</div><div style="font:900 39px/1.15 Nunito, sans-serif;color:#673b51;letter-spacing:-1.2px">Every order deserves<br>a thoughtful reply.</div><p style="color:#89677a;font-size:14px;line-height:1.75;margin-top:13px;max-width:490px">Bring customer questions and store policies together, so your team can respond with clarity and care.</p><div class="hero-note">✦ Made for the moments after checkout</div></div>''', unsafe_allow_html=True)
with art_col:
    st.image(str(HERE / "parcel.svg"), use_container_width=True)

st.markdown('''<div class="microheading">At a glance</div><div class="legendbar">
<div class="legend"><i>📦</i>Order enquiries</div>
<div class="legend"><i>🚚</i>Shipping & delivery</div>
<div class="legend"><i>↩️</i>Returns & refunds</div>
<div class="legend"><i>📖</i>Store policies</div>
<div class="legend"><i>✉️</i>Reply drafts</div>
</div>''', unsafe_allow_html=True)

completed = 3 if st.session_state.get("result") else 0
flow = [("01", "Email received"), ("02", "Policy matched"), ("03", "Draft prepared")]
flow_html = []
for i, (num, label) in enumerate(flow):
    flow_html.append(f'<div class="flowstep"><span class="stepdot {"done" if completed > i else ""}">{"✓" if completed > i else num}</span><span class="steptxt">{label}</span></div>')
    if i < 2:
        flow_html.append('<span class="steparrow">→</span>')
st.markdown('<div class="flowline">' + ''.join(flow_html) + '</div>', unsafe_allow_html=True)

left, right = st.columns([0.95, 1.05], gap="large")
with left:
    st.markdown('<div class="area-title">📨 Incoming enquiry</div><div class="area-desc">Add customer details to prepare a new support response.</div>', unsafe_allow_html=True)
    a, b = st.columns(2)
    with a:
        if st.button("🚚 Try delivery case", use_container_width=True):
            st.session_state.update({"customer_name": EXAMPLES["delivery"]["name"], "customer_email": EXAMPLES["delivery"]["email"], "customer_body": EXAMPLES["delivery"]["body"]})
            st.session_state.pop("result", None)
            st.rerun()
    with b:
        if st.button("↩️ Try return case", use_container_width=True):
            st.session_state.update({"customer_name": EXAMPLES["wrong_item"]["name"], "customer_email": EXAMPLES["wrong_item"]["email"], "customer_body": EXAMPLES["wrong_item"]["body"]})
            st.session_state.pop("result", None)
            st.rerun()
    with st.form("customer_form"):
        customer_name = st.text_input("Customer name", key="customer_name")
        customer_email = st.text_input("Customer email", key="customer_email")
        customer_body = st.text_area("Customer message", height=240, key="customer_body")
        submitted = st.form_submit_button("Prepare reply   →", use_container_width=True)
    st.markdown('<div class="explain">🔒 The workspace creates a draft only. An agent reviews the content before contacting the customer.</div>', unsafe_allow_html=True)

if submitted:
    st.session_state.pop("result", None)
    st.session_state.pop("draft_reply_box", None)
    if not api_key.strip():
        st.error("Enter your HKUST GenAI Primary Key using the settings menu on the left.")
    elif not customer_name.strip() or not customer_body.strip():
        st.error("Please provide a customer name and a message.")
    elif not all(path.is_file() for path in DOCS.values()):
        st.error("Knowledge base documents are missing. Keep the data folder beside app.py.")
    else:
        try:
            client = client_for(api_key.strip(), base_url.strip())
            with st.status("Preparing your reply", expanded=True) as status:
                progress = st.progress(0, text="Reading customer enquiry")
                category = classify_email(client, chat_model.strip(), customer_body)
                progress.progress(34, text="Enquiry classified · finding store policies")
                passages = retrieve(client, category, customer_body)
                progress.progress(68, text="Relevant policies found · drafting reply")
                reply = draft_reply(client, chat_model.strip(), customer_name.strip(), customer_email.strip(), customer_body, category, passages)
                progress.progress(100, text="Draft complete · ready for agent review")
                status.update(label="Reply ready for review", state="complete", expanded=False)
            st.session_state["result"] = {"category": category, "passages": passages, "reply": reply, "name": customer_name, "email": customer_email, "body": customer_body}
            st.rerun()
        except Exception as exc:
            st.error(f"Could not prepare reply: {type(exc).__name__}: {exc}")
            st.caption("Check your API key, endpoint, model selection, and network connection.")

with right:
    st.markdown('<div class="area-title">✉️ Reply workspace</div><div class="area-desc">The prepared response and the policies behind it.</div>', unsafe_allow_html=True)
    r = st.session_state.get("result")
    if r:
        icon = "🚚" if r["category"] == "Shipping / Delivery" else "↩️"
        st.markdown(f'<div class="case-card"><div class="case-head"><span class="tag">{icon} {escape(r["category"])}</span><span class="tag green">✓ Draft ready</span></div><div class="case-meta">CUSTOMER &nbsp; {escape(r["name"])} &nbsp; • &nbsp; {escape(r["email"] or "No email provided")}</div><div class="stagefill"><div style="width:100%"></div></div></div>', unsafe_allow_html=True)
        draft_tab, evidence_tab, summary_tab = st.tabs(["✉️ Reply draft", "📖 Policy sources", "📦 Case notes"])
        with draft_tab:
            st.markdown('<div class="result-kicker">Ready for human review</div>', unsafe_allow_html=True)
            st.text_area("Draft response — editable", value=r["reply"], key="draft_reply_box", height=305)
            st.download_button("↓ Download email (.txt)", data=st.session_state.get("draft_reply_box", r["reply"]), file_name="acme_reply.txt", mime="text/plain", use_container_width=True)
            st.caption("No email is sent automatically.")
        with evidence_tab:
            st.caption("Excerpts selected from the matching knowledge collection. Similarity indicates semantic closeness, not the probability that an answer is correct.")
            for i, item in enumerate(r["passages"], 1):
                st.markdown(f'<div class="evidence"><div class="case-head"><div><div class="evidence-title">{i:02d} · {escape(item["section"])}</div><div class="evidence-small">{escape(item["source"])}</div></div><span class="evidence-score">Similarity {item["score"]:.3f}</span></div></div>', unsafe_allow_html=True)
                with st.expander("View policy excerpt", expanded=(i == 1)):
                    st.write(item["text"])
        with summary_tab:
            st.markdown("**Case overview**")
            st.write(f"**Customer:** {r['name']}")
            st.write(f"**Category:** {r['category']}")
            st.write(f"**Matched sources:** {len(r['passages'])}")
            st.markdown("**Original enquiry**")
            st.write(r["body"])
            st.info("Next step: Review the drafted email and handle the case through your normal support process.")
    else:
        st.markdown('''<div class="case-card" style="padding:28px 26px"><div style="font-size:32px;margin-bottom:9px">💌</div><div style="font-family:Nunito,sans-serif;font-size:20px;font-weight:900;color:#72485b;margin-bottom:8px">A little care goes a long way.</div><div style="font-size:13px;color:#9c7788;line-height:1.7">Select an example or enter a customer message. Your reply draft, policy references and case notes will appear here.</div></div>''', unsafe_allow_html=True)
        st.markdown("**How your case is handled**")
        st.markdown("📨 **Receive** the customer's message  \n📖 **Look up** the matching policy  \n✉️ **Prepare** a response for staff review")

st.markdown('<div class="footerline">ACME · ORDER CARE &nbsp; | &nbsp; Support drafts informed by company policy · Human-reviewed before sending</div>', unsafe_allow_html=True)
