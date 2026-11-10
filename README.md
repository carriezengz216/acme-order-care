# Acme Order Care — Customer Service RAG (Pink Studio Edition)

A Python + Streamlit reconstruction of the team's two-branch Dify customer service workflow. The application takes a customer name, email address, and message; classifies the enquiry as shipping/delivery or refund/return/exchange; retrieves semantically relevant passages from the corresponding original DOCX knowledge base; and prepares a human-reviewed customer service reply.

## Run on macOS

Unzip the project. Keep `app.py`, `data/`, and `assets/` together. If you are replacing a previous version, copying just `app.py` **and** the `assets/` folder into your existing project also works.

In Terminal:

```bash
conda activate acme_rag
cd ~/Desktop/Customer_Service_RAG_Studio
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

If you don't already have `acme_rag`, create it first with `conda create -n acme_rag python=3.11 -y`, then activate it. If you renamed the unzipped folder, update the `cd` command accordingly. The application normally appears at http://localhost:8501.

## Connect the API

Open the left sidebar to enter your **HKUST GenAI Primary Key**. It is entered at runtime and never hardcoded in the project. The default API base URL is `https://hkust.azure-api.net/hkust-genai/v1/`, with `text-embedding-ada-002` for embeddings and `gpt-4.1-mini` for classification and drafting. The chat model and URL can be edited under Advanced API settings if necessary.

Don't expose the key in a screen recording or upload it with your project files. This app doesn't send emails, change customer orders, or provide actual parcel tracking updates.

## Two ready-made test cases

Use the buttons above the form:

- **Try delivery case**: Megan asks about an order delay and stale tracking. Expected label: Shipping / Delivery.
- **Try return case**: Marcus received a chair instead of the ordered SmartDesk Pro 2. Expected label: Refund / Return / Exchange.

After pressing **Prepare reply**, the app shows the routed case, editable email draft, retrieved policy passages with cosine similarities, and case notes.

## Design and workflow choices

- Brand-oriented pink/white order-care interface with an original parcel illustration in `assets/parcel.svg`, purposeful icons, and understated status badges.
- Three-stage process indicator for email processing—not a real logistics tracking timeline.
- Progress updates during classification, retrieval, and drafting.
- Source evidence is separated from the reply draft to keep the workspace usable and the RAG mechanism verifiable.
- The API connection and RAG logic are retained from the previous implementation. Document vectors are kept in Streamlit session state and created on demand.

## Files

- `app.py`: Streamlit UI and RAG pipeline.
- `data/Shipping_Delivery_Issues.docx`: team's original delivery knowledge base.
- `data/Refund_Return_Exchange.docx`: team's original refunds knowledge base.
- `assets/parcel.svg`: original parcel illustration drawn for this UI.
- `requirements.txt`: Python dependencies.

## Suggested 2-minute video

0:00–0:15: Introduce Acme Order Care and explain the customer support use case.  
0:15–0:35: Show the two policy knowledge collections and the three-stage workflow.  
0:35–1:25: Use a delivery case and show classification, draft and policy sources.  
1:25–1:50: Use the return case to demonstrate the second knowledge branch.  
1:50–2:00: Summarise that the system drafts grounded responses for human review.

The presenter must remain visible on camera while speaking, as required by the instructor. Test the app before recording. A real API connection is needed for full end-to-end operation.
