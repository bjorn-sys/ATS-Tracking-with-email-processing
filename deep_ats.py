import re
import io
import json
import smtplib
import imaplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from collections import Counter
import streamlit as st

# Optional libraries for file parsing
try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    import docx2txt
except Exception:
    docx2txt = None

# --------------------------- Email Configuration ---------------------------
# These would typically be stored as environment variables or in a config file
EMAIL_ADDRESS = "your_email@example.com"
EMAIL_PASSWORD = "your_app_password"
IMAP_SERVER = "imap.example.com"  # e.g., imap.gmail.com
SMTP_SERVER = "smtp.example.com"  # e.g., smtp.gmail.com

# --------------------------- Helpers ---------------------------

CONTACT_PATTERNS = {
    "email": re.compile(r"[\w\.-]+@[\w\.-]+\.[a-zA-Z]{2,6}"),
    "phone": re.compile(r"(\+\d{1,3}[\s-])?(\(?\d{3}\)?[\s-]?)?\d{3}[\s-]?\d{4}"),
    "linkedin": re.compile(r"linkedin\.com/[A-Za-z0-9_\-/%]+"),
}

SECTION_KEYWORDS = {
    "contact": ["contact", "email", "phone", "mobile", "linkedin"],
    "summary": ["summary", "profile", "about me", "overview"],
    "skills": ["skills", "technical skills", "tools", "technologies"],
    "experience": ["experience", "work experience", "employment", "roles"],
    "education": ["education", "academic", "degree", "bachelor", "master", "university"],
}

STOPWORDS = set([
    "the", "and", "a", "an", "of", "to", "in", "for", "on", "with", "as",
    "is", "are", "by", "that", "this", "be", "at", "or", "from", "it",
])

# --------------------------- Email Functions ---------------------------

def connect_to_email():
    """Connect to IMAP server and return the connection"""
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        return mail
    except Exception as e:
        st.error(f"Failed to connect to email: {str(e)}")
        return None

def fetch_emails():
    """Fetch emails with attachments from inbox"""
    mail = connect_to_email()
    if not mail:
        return []
    
    try:
        mail.select("inbox")
        result, data = mail.search(None, 'ALL')
        email_ids = data[0].split()
        
        emails = []
        for email_id in email_ids:
            result, data = mail.fetch(email_id, "(RFC822)")
            raw_email = data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # Check if email has attachments
            has_attachment = any(part.get_content_disposition() == 'attachment' for part in msg.walk())
            
            if has_attachment:
                email_info = {
                    'id': email_id,
                    'subject': msg['subject'] or 'No Subject',
                    'from': msg['from'],
                    'date': msg['date'],
                    'attachments': []
                }
                
                # Extract attachments
                for part in msg.walk():
                    if part.get_content_disposition() == 'attachment':
                        filename = part.get_filename()
                        if filename:
                            file_data = part.get_payload(decode=True)
                            email_info['attachments'].append({
                                'filename': filename,
                                'data': file_data
                            })
                
                emails.append(email_info)
        
        return emails
    except Exception as e:
        st.error(f"Error fetching emails: {str(e)}")
        return []
    finally:
        if mail:
            mail.logout()

def move_email_to_folder(mail, email_id, folder_name="Screening"):
    """Move email to specified folder"""
    try:
        # Create folder if it doesn't exist
        mail.create(folder_name)
    except:
        pass  # Folder might already exist
    
    try:
        # Copy email to new folder
        mail.copy(email_id, folder_name)
        # Mark original for deletion
        mail.store(email_id, '+FLAGS', '\\Deleted')
        mail.expunge()
        return True
    except Exception as e:
        st.error(f"Error moving email: {str(e)}")
        return False

def send_email(to_address, subject, body):
    """Send an email notification"""
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = to_address
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain'))
        
        server = smtplib.SMTP_SSL(SMTP_SERVER, 465)
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Error sending email: {str(e)}")
        return False

# --------------------------- Resume Processing Functions ---------------------------

def extract_text_from_pdf(file_bytes):
    if pdfplumber is None:
        return ""
    text = []
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text.append(page_text)
    except Exception:
        return ""
    return "\n".join(text)

def extract_text_from_docx(file_bytes):
    if docx2txt is None:
        return ""
    try:
        # docx2txt works with filesystem paths; write to temp buffer
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        text = docx2txt.process(tmp_path) or ""
    except Exception:
        return ""
    return text

def extract_text_from_bytes(file_bytes, filename):
    name = filename.lower()
    if name.endswith(".pdf"):
        text = extract_text_from_pdf(file_bytes)
    elif name.endswith(".docx") or name.endswith("doc"):
        text = extract_text_from_docx(file_bytes)
    else:
        try:
            text = file_bytes.decode("utf-8")
        except Exception:
            try:
                text = file_bytes.decode("latin-1")
            except Exception:
                text = ""
    # fallback empty string -> try simple binary->string
    return text or ""

def normalize_text(text):
    text = re.sub(r"[^\w\s-]", " ", text)
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text

def detect_sections(text):
    found = {}
    ln = text.lower()
    for sec, keywords in SECTION_KEYWORDS.items():
        found[sec] = any(k in ln for k in keywords)
    return found

def detect_contact(text):
    contact = {}
    for k, pat in CONTACT_PATTERNS.items():
        contact[k] = bool(pat.search(text))
    return contact

def simple_keyword_extract(text, top_k=30):
    text = normalize_text(text)
    tokens = [t for t in text.split() if t not in STOPWORDS and len(t) > 2]
    counts = Counter(tokens)
    most = [w for w, _ in counts.most_common(top_k)]
    return most

def keyword_match_score(resume_text, jd_text, top_k=30):
    if not jd_text.strip():
        return 0.0, []
    jd_keywords = simple_keyword_extract(jd_text, top_k=top_k)
    resume_tokens = set(normalize_text(resume_text).split())
    matched = [w for w in jd_keywords if w in resume_tokens]
    score = len(matched) / max(1, len(jd_keywords))
    return float(score * 100), matched

def compute_length_score(resume_text, min_words=200, ideal_words=500, max_words=1200):
    words = len(normalize_text(resume_text).split())
    if words == 0:
        return 0.0, words
    if min_words <= words <= max_words:
        # closer to ideal -> higher score
        diff = abs(words - ideal_words)
        score = max(0, 100 - (diff / ideal_words) * 100)
    else:
        score = 40.0 * (min(1.0, words / max_words))
    return float(score), words

def bullet_score(resume_text, ideal_bullets=10):
    bullets = resume_text.count("\n-") + resume_text.count("\n•") + resume_text.count("\n*")
    if bullets >= ideal_bullets:
        return 100.0, bullets
    else:
        return float((bullets / ideal_bullets) * 100), bullets

def aggregate_score(section_found, contact, keyword_pct, length_pct, bullet_pct, weights=None):
    # section_found is dict, contact is dict
    sections_score = (sum(section_found.values()) / len(section_found)) * 100
    contact_score = (1.0 if any(contact.values()) else 0.0) * 100
    if weights is None:
        weights = {
            "sections": 0.20,
            "contact": 0.10,
            "keywords": 0.35,
            "length": 0.15,
            "bullets": 0.20,
        }
    total = (
        sections_score * weights["sections"]
        + contact_score * weights["contact"]
        + keyword_pct * weights["keywords"]
        + length_pct * weights["length"]
        + bullet_pct * weights["bullets"]
    )
    return round(float(total), 1)

def generate_suggestions(section_found, contact, matched_keywords, missing_keywords, words_count, bullets_count):
    suggestions = []
    # Sections
    missing_sections = [s for s, v in section_found.items() if not v]
    if missing_sections:
        suggestions.append(f"Add or expand these sections: {', '.join(missing_sections)}.")
    # Contact
    if not any(contact.values()):
        suggestions.append("Include clear contact details (email, phone, and LinkedIn).")
    # Keywords
    if matched_keywords:
        suggestions.append(f"You matched {len(matched_keywords)} important keywords from the job description. Keep them natural and contextual.")
    if missing_keywords:
        suggestions.append(f"Consider adding these keywords where relevant: {', '.join(missing_keywords[:10])}.")
    # Length
    if words_count < 200:
        suggestions.append("Your resume is quite short. Aim for 1 page (about 300–500 words) unless you have extensive experience.")
    elif words_count > 1200:
        suggestions.append("Your resume is long. Aim to be concise — 1–2 pages is typical.")
    # Bullets
    if bullets_count < 5:
        suggestions.append("Use bullet points under roles to highlight achievements and metrics.")
    # General
    suggestions.append("Use action verbs, quantifiable results, and align phrasing with the job description." )
    return suggestions

def process_resume(file_bytes, filename, job_description="", top_k=25, weights=None):
    """Process a resume and return the score and analysis"""
    resume_text = extract_text_from_bytes(file_bytes, filename)
    if not resume_text.strip():
        return None
    
    # Normalize & detect
    norm_resume = normalize_text(resume_text)
    sections = detect_sections(resume_text)
    contact = detect_contact(resume_text)
    keyword_pct, matched = keyword_match_score(resume_text, job_description or "", top_k=top_k)
    # missing keywords from jd
    jd_keywords = simple_keyword_extract(job_description, top_k=top_k) if job_description.strip() else []
    missing_kw = [k for k in jd_keywords if k not in normalize_text(resume_text).split()]
    length_pct, words_count = compute_length_score(resume_text)
    bullets_pct, bullets_count = bullet_score(resume_text)

    overall = aggregate_score(sections, contact, keyword_pct, length_pct, bullets_pct, weights=weights)
    
    return {
        'score': overall,
        'text': resume_text,
        'sections': sections,
        'contact': contact,
        'keyword_pct': keyword_pct,
        'matched_keywords': matched,
        'missing_keywords': missing_kw,
        'words_count': words_count,
        'bullets_count': bullets_count
    }

# --------------------------- Streamlit UI ---------------------------

st.set_page_config(page_title="ATS Resume Scorer", page_icon="📄", layout="wide")
st.title("ATS Resume Scorer with Email Processing")
st.caption("Automatically fetch resumes from email, screen them, and move qualified candidates to screening folder.")

# Configuration section
with st.sidebar:
    st.header("Configuration")
    st.text_input("Email Address", value=EMAIL_ADDRESS, key="email_address")
    st.text_input("Email Password", type="password", value=EMAIL_PASSWORD, key="email_password")
    st.text_input("IMAP Server", value=IMAP_SERVER, key="imap_server")
    st.text_input("SMTP Server", value=SMTP_SERVER, key="smtp_server")
    
    st.header("Scoring Settings")
    top_k = st.slider("Number of keywords to extract from JD", min_value=5, max_value=60, value=25)
    score_threshold = st.slider("Minimum passing score", min_value=0, max_value=100, value=80)
    weights_input = st.text_area("Optional custom weights (json)", value=json.dumps({
        "sections": 0.20,
        "contact": 0.10,
        "keywords": 0.35,
        "length": 0.15,
        "bullets": 0.20
    }, indent=2), height=150)
    
    job_description = st.text_area("Paste job description", height=200, help="This will be used to score all resumes")

# Main content
tab1, tab2, tab3 = st.tabs(["Email Processing", "Manual Upload", "Results"])

# Parse weights
try:
    custom_weights = json.loads(weights_input)
except Exception:
    st.warning("Invalid JSON for weights — using defaults.")
    custom_weights = None

with tab1:
    st.header("Process Emails with Resumes")
    if st.button("Fetch and Process Emails"):
        with st.spinner("Fetching emails..."):
            emails = fetch_emails()
        
        if not emails:
            st.info("No emails with attachments found.")
        else:
            st.write(f"Found {len(emails)} emails with attachments")
            
            qualified = []
            disqualified = []
            
            for email_data in emails:
                st.subheader(f"From: {email_data['from']} - {email_data['subject']}")
                
                for attachment in email_data['attachments']:
                    filename = attachment['filename']
                    if any(filename.lower().endswith(ext) for ext in ['.pdf', '.docx', '.doc', '.txt']):
                        st.write(f"Processing: {filename}")
                        
                        result = process_resume(
                            attachment['data'], 
                            filename, 
                            job_description, 
                            top_k, 
                            custom_weights
                        )
                        
                        if result:
                            score = result['score']
                            st.write(f"Score: {score}/100")
                            
                            if score >= score_threshold:
                                st.success("✅ Qualified - Moving to screening folder")
                                qualified.append({
                                    'email': email_data,
                                    'attachment': attachment,
                                    'score': score,
                                    'result': result
                                })
                            else:
                                st.error("❌ Disqualified - Score below threshold")
                                disqualified.append({
                                    'email': email_data,
                                    'attachment': attachment,
                                    'score': score,
                                    'result': result
                                })
                        else:
                            st.warning("Could not process this attachment")
                    else:
                        st.write(f"Skipping non-resume file: {filename}")
            
            # Move emails to appropriate folders
            if qualified:
                mail = connect_to_email()
                if mail:
                    for item in qualified:
                        move_email_to_folder(mail, item['email']['id'], "Screening")
                    mail.logout()
                
                st.success(f"Moved {len(qualified)} qualified resumes to screening folder")
            
            # Store results in session state for display in Results tab
            st.session_state.qualified = qualified
            st.session_state.disqualified = disqualified

with tab2:
    st.header("Manual Resume Upload")
    uploaded_file = st.file_uploader("Upload resume (PDF, DOCX, TXT)", type=["pdf", "docx", "doc", "txt"]) 
    
    if st.button("Analyze Resume"):
        if not uploaded_file:
            st.error("Please upload a resume file first.")
        else:
            with st.spinner("Processing resume..."):
                result = process_resume(
                    uploaded_file.read(), 
                    uploaded_file.name, 
                    job_description, 
                    top_k, 
                    custom_weights
                )
            
            if not result:
                st.error("Could not extract any text from the file.")
            else:
                score = result['score']
                st.metric("ATS Score", f"{score} / 100")
                
                if score >= score_threshold:
                    st.success("✅ This resume qualifies for screening")
                else:
                    st.error("❌ This resume does not meet the minimum score")
                
                # Display details
                st.subheader("Breakdown")
                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Section presence**")
                    for s, v in result['sections'].items():
                        st.write(f"{s.title()}: {'✔️' if v else '❌'}")
                    st.write("")
                    st.write("**Contact detected**")
                    for k, v in result['contact'].items():
                        st