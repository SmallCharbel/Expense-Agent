from flask import Flask, render_template, request, redirect, url_for, flash
import os
import re
from datetime import datetime, timedelta
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-in-production'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.config['ALLOWED_EXTENSIONS_PDF'] = {'pdf'}
app.config['ALLOWED_EXTENSIONS_EXCEL'] = {'xlsx', 'xls'}

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def allowed_file(filename, file_type):
    if file_type == 'pdf':
        return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS_PDF']
    elif file_type == 'excel':
        return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS_EXCEL']
    return False

def derive_name_from_excel_filename(filename):
    # This function attempts to extract a person's name from the Excel file name.
    # It's a heuristic and may need adjustment for different naming conventions.

    # 1. Remove the timestamp prefix if it exists (e.g., "20251021_122023_")
    name = re.sub(r"^\d{8}_\d{6}_", "", filename)

    # 2. Remove the file extension
    name = re.sub(r"\.(xlsx|xls)$", "", name, flags=re.IGNORECASE)

    # 3. Remove common keywords like "Expenses" or "Report"
    name = re.sub(r"(?i)\s*\b(Expenses|Report)\b\s*", "", name)

    # 4. Insert a space before a capital letter in a camelCase-like name (e.g., "JanetS" -> "Janet S")
    name = re.sub(r"([a-z])([A-Z])", r"\1 \2", name)

    # 5. Replace common separators with spaces
    name = name.replace("_", " ").replace("-", " ")

    # 6. Remove any remaining non-alphanumeric characters except spaces (e.g., commas)
    name = re.sub(r"[^a-zA-Z0-9\s]", "", name)

    # 7. Trim whitespace from the ends and handle multiple spaces
    return " ".join(name.split())

def read_pdf_text(pdf_path):
    try:
        import fitz
        doc = fitz.open(pdf_path)
        text = "\n".join(page.get_text("text") for page in doc)
        doc.close()
    except Exception as e:
        print(f"PyMuPDF failed: {e}, trying basic read")
        with open(pdf_path, 'rb') as f:
            text = f.read().decode('utf-8', errors='ignore')
    return re.sub(r" +", " ", re.sub(r"[\t\r\u00a0]+", " ", text))

def parse_statement_period(pdf_text):
    m = re.search(r"Closing Date\s*(\d{2}/\d{2}/\d{2})", pdf_text)
    if not m:
        raise ValueError("Closing Date not found in PDF")
    end = datetime.strptime(m.group(1), "%m/%d/%y").date()
    return end - timedelta(days=30), end

def find_cardholder_sections(pdf_text):
    sections = [(m.group(1).strip(), m.start()) for m in re.finditer(r"([A-Z][A-Z ]{1,80}?)\s+Card Ending", pdf_text, re.MULTILINE)]
    return sorted(sections, key=lambda x: x[1])

def select_target_section_name(target, available):
    # Normalize the target name derived from the filename
    target_norm = target.upper()
    target_tokens = set(target_norm.split())

    # Normalize available names from the PDF and try for an exact match first
    available_norm = []
    for name, pos in available:
        name_norm = name.strip().upper()
        if name_norm == target_norm:
            # Exact match is the best possible outcome
            return name
        available_norm.append({'original': name, 'normalized': name_norm, 'pos': pos})

    # If no exact match, proceed with a scoring mechanism
    candidates = []
    for item in available_norm:
        name_norm = item['normalized']
        name_tokens = set(name_norm.split())

        # Scoring logic
        score = 0
        # 1. Reward for full token matches
        score += len(target_tokens.intersection(name_tokens)) * 10

        # 2. Reward for being a substring (e.g., 'JOSE' in 'JOSEPH')
        if target_norm in name_norm or name_norm in target_norm:
            score += 5

        # 3. Reward for initial matches (e.g., 'JanetS' vs 'JANET S')
        # This helps with cases where last name initial is attached.
        for t_token in target_tokens:
            for n_token in name_tokens:
                if t_token.startswith(n_token) or n_token.startswith(t_token):
                    score += 2
        
        if score > 0:
            candidates.append((score, item['original']))

    # Return the best candidate, or the first available name as a fallback
    if candidates:
        return max(candidates, key=lambda x: x[0])[1]
    elif available:
        # Fallback to the first name found if no match at all
        return available[0][0]
    else:
        # No names found in PDF
        return None

def extract_transactions_for_name(pdf_text, target_name, sections):
    target_pos = next((pos for name, pos in sections if name == target_name), None)
    if not target_pos:
        return []

    # 1. Find the start of the next cardholder's section
    positions = sorted([pos for _, pos in sections])
    next_cardholder_pos = next((p for p in positions if p > target_pos), None)

    # 2. Find the start of a potential summary section after the current cardholder
    # These keywords mark the end of transactions. We look for lines starting with them.
    summary_keywords = [
        "FEES", "INTEREST CHARGED", "ACCOUNT SUMMARY", 
        "PAYMENT INFORMATION", "LATE FEE", "TOTAL CREDIT", "TOTAL DEBIT",
        "ABOUT TRAILING INTEREST"
    ]
    
    text_after_target = pdf_text[target_pos:]
    summary_pos = -1

    # Search for the first occurrence of any of these keywords at the beginning of a line.
    for keyword in summary_keywords:
        # The regex looks for the keyword at the start of a line, ignoring leading whitespace.
        match = re.search(r"^\s*" + re.escape(keyword), text_after_target, re.IGNORECASE | re.MULTILINE)
        if match:
            # Adjust position to be relative to the full pdf_text
            actual_pos = target_pos + match.start()
            if summary_pos == -1 or actual_pos < summary_pos:
                summary_pos = actual_pos

    # 3. Determine the end of the section (next_pos)
    # It's the earliest of: next cardholder's start, summary section start, or end of document.
    end_positions = [pos for pos in [next_cardholder_pos, summary_pos] if pos is not None and pos > target_pos]
    
    if end_positions:
        next_pos = min(end_positions)
    else:
        next_pos = len(pdf_text) # Fallback to end of document

    section_text = pdf_text[target_pos:next_pos]
    
    transactions = []
    for chunk in re.split(r"[\u25CA\u2B27\u2B25\u25C6\u25C7\u29EB\u2B26\u2B29|\u29eb]+", section_text):
        ch = ' '.join(chunk.split())
        m_date = re.search(r"(\d{2}/\d{2}/\d{2})", ch)
        m_amt = list(re.finditer(r"\$([0-9,]+\.[0-9]{2})", ch))
        if m_date and m_amt:
            merchant = ch[m_date.end():m_amt[-1].start()].strip()
            if not merchant.upper().startswith('CARD ENDING'):
                transactions.append({
                    'date': m_date.group(1),
                    'merchant': merchant,
                    'amount': float(m_amt[-1].group(1).replace(',', ''))
                })
    return transactions

def read_excel_transactions(xlsx_path, start_date, end_date):
    all_txns = []
    try:
        xls = pd.ExcelFile(xlsx_path)
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(xlsx_path, sheet_name=sheet)
                date_col = next((c for c in df.columns if 'date' in str(c).lower()), None)
                amt_col = next((c for c in df.columns if any(x in str(c).lower() for x in ['amount', 'total', 'charge'])), None)
                if not date_col or not amt_col:
                    continue
                
                df[amt_col] = pd.to_numeric(df[amt_col].astype(str).str.replace('[$,]', '', regex=True), errors='coerce')
                df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                mask = (df[date_col].dt.date >= start_date) & (df[date_col].dt.date <= end_date)
                
                for _, row in df[mask].dropna(subset=[date_col, amt_col]).iterrows():
                    all_txns.append({'date': row[date_col].date().strftime('%m/%d/%y'), 'amount': float(row[amt_col])})
            except Exception as e:
                print(f"Error reading sheet {sheet}: {e}")
                continue
    except Exception as e:
        print(f"Error reading Excel file: {e}")
    return all_txns

def match_transactions(pdf_txns, excel_txns):
    available = [t['amount'] for t in excel_txns]
    missing = []
    matched = 0
    for t in pdf_txns:
        idx = next((i for i, amt in enumerate(available) if amt - 0.50 <= t['amount'] <= amt + 0.50), None)
        if idx is not None:
            matched += 1
            available.pop(idx)
        else:
            missing.append(t)
    return matched, missing

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/compare', methods=['POST'])
def compare():
    # Check if files are in request
    if 'pdf' not in request.files or 'excel' not in request.files:
        flash('Both PDF and Excel files are required!', 'error')
        return redirect(url_for('index'))
    
    pdf_file = request.files['pdf']
    excel_file = request.files['excel']
    
    # Check if files are selected
    if pdf_file.filename == '' or excel_file.filename == '':
        flash('Please select both files!', 'error')
        return redirect(url_for('index'))
    
    # Validate file types
    if not allowed_file(pdf_file.filename, 'pdf'):
        flash('PDF file must be a .pdf file!', 'error')
        return redirect(url_for('index'))
    
    if not allowed_file(excel_file.filename, 'excel'):
        flash('Excel file must be .xlsx or .xls!', 'error')
        return redirect(url_for('index'))
    
    # Secure filenames
    pdf_filename = secure_filename(pdf_file.filename)
    excel_filename = secure_filename(excel_file.filename)
    
    # Add timestamp to avoid conflicts
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    pdf_filename = f"{timestamp}_{pdf_filename}"
    excel_filename = f"{timestamp}_{excel_filename}"
    
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
    
    try:
        # Save files
        pdf_file.save(pdf_path)
        excel_file.save(excel_path)
        
        # Process files
        person_name = derive_name_from_excel_filename(excel_file.filename)
        pdf_text = read_pdf_text(pdf_path)
        start_date, end_date = parse_statement_period(pdf_text)
        
        sections = find_cardholder_sections(pdf_text)
        if not sections:
            raise ValueError("No cardholders found in PDF")
        
        selected = select_target_section_name(person_name, sections)
        if not selected:
            raise ValueError(f"Could not match '{person_name}' to any cardholder")
        
        pdf_txns = extract_transactions_for_name(pdf_text, selected, sections)
        excel_txns = read_excel_transactions(excel_path, start_date, end_date)
        matched, missing = match_transactions(pdf_txns, excel_txns)
        
        pdf_total = sum(t['amount'] for t in pdf_txns)
        excel_total = sum(t['amount'] for t in excel_txns)
        missing_total = sum(t['amount'] for t in missing)
        
        results = {
            'cardholder': selected,
            'period_start': start_date.strftime('%m/%d/%y'),
            'period_end': end_date.strftime('%m/%d/%y'),
            'pdf_count': len(pdf_txns),
            'pdf_total': f"{pdf_total:,.2f}",
            'excel_count': len(excel_txns),
            'excel_total': f"{excel_total:,.2f}",
            'matched': matched,
            'missing_count': len(missing),
            'missing_total': f"{missing_total:,.2f}",
            'missing_transactions': missing
        }
        
        # Clean up files
        try:
            os.remove(pdf_path)
            os.remove(excel_path)
        except:
            pass
        
        return render_template('results.html', results=results)
    
    except Exception as e:
        # Clean up files on error
        try:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            if os.path.exists(excel_path):
                os.remove(excel_path)
        except:
            pass
        
        error_message = str(e)
        print(f"Error: {error_message}")
        return render_template('error.html', error=error_message)

@app.errorhandler(413)
def too_large(e):
    return render_template('error.html', error="File too large! Maximum size is 50MB."), 413

if __name__ == '__main__':
    print("=" * 60)
    print("🚀 Expense Comparator Starting...")
    print("=" * 60)
    print(f"📁 Upload folder: {os.path.abspath(app.config['UPLOAD_FOLDER'])}")
    print(f"📊 Max file size: 50MB")
    print(f"🌐 Open: http://localhost:5000")
    print("=" * 60)
    app.run(debug=True, port=5000, host='0.0.0.0')