# PO_extractor_Lingbo_group
import fitz  # PyMuPDF
import re
import zipfile
import os
import shutil
import pandas as pd
from pathlib import Path

# -----------------------------
# Function to extract required info
# -----------------------------
def extract_po_info(pdf_path: str) -> list:
    rows = []
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            blocks = page.get_text("blocks")
            blocks.sort(key=lambda b: (b[1], b[0]))
            for b in blocks:
                full_text += b[4] + "\n"
        doc.close()

        lines = [l.strip() for l in full_text.splitlines() if l.strip()]

        # -------------------
        # Document Number
        # -------------------
        first_int = re.search(r"\d+", full_text)
        document_number = f"PO{first_int.group(0)}" if first_int else ""

        # -------------------
        # Header info
        # -------------------
        po_issue_date, payment_term, ship_via, fob = "", "", "", ""
        for i, line in enumerate(lines):
            if not po_issue_date and re.search(r"\d{1,2}/\d{1,2}/\d{4}", line):
                po_issue_date = re.search(r"\d{1,2}/\d{1,2}/\d{4}", line).group(0)

            net_match = re.search(r"(Net\s*\d+)", line, re.IGNORECASE)
            if net_match:
                payment_term = net_match.group(1)
                fob = line[:net_match.start()].strip()
                if i-1 >= 0:
                    upper_line = lines[i-1].strip()
                    if not re.match(r"^\d{1,2}/\d{1,2}/\d{2,4}", upper_line) \
                       and not re.match(r"^\d+$", upper_line):
                        ship_via = upper_line
                break

        # -------------------
        # Buyer, REQ# and Requisitioner (dynamic logic)
        # -------------------
        buyer, req_num, requisitioner = "", "", ""
        for i, line in enumerate(lines):
            if line.strip().upper() == "ITEM":
                # Requisitioner: 1 line above ITEM
                candidate_req_name = lines[i-1].strip() if i-1 >= 0 else ""
                if candidate_req_name and not re.match(r"\d{1,2}/\d{1,2}/\d{2,4}", candidate_req_name) \
                   and not re.match(r"^\d+$", candidate_req_name) \
                   and not re.search(r"Net\s*\d+", candidate_req_name, re.IGNORECASE):
                    requisitioner = candidate_req_name
                else:
                    requisitioner = ""

                # REQ#: 2 lines above ITEM
                candidate_req = lines[i-2].strip() if i-2 >= 0 else ""
                if re.match(r"^\d+$", candidate_req):
                    req_num = candidate_req
                else:
                    req_num = ""

                # Buyer: dynamic based on other fields
                if requisitioner == "" and req_num == "":
                    candidate_buyer = lines[i-1].strip() if i-1 >= 0 else ""  # 1 line above
                elif requisitioner == "":
                    candidate_buyer = lines[i-2].strip() if i-2 >= 0 else ""  # 2 lines above
                else:
                    candidate_buyer = lines[i-3].strip() if i-3 >= 0 else ""  # 3 lines above

                buyer = candidate_buyer if re.match(r"^\d+$", candidate_buyer) else ""
                break

        # -------------------
        # Parse line-item blocks
        # -------------------
        blocks = []
        in_block_section = False
        current_block = []
        skip_header = False

        for i, line in enumerate(lines):
            if not in_block_section:
                if i < len(lines)-1 and lines[i] == "Unit Cost" and lines[i+1] == "Extended Cost":
                    in_block_section = True
                    skip_header = True
                continue

            if skip_header and line == "Extended Cost":
                skip_header = False
                continue

            if line.strip().upper() == "TOTAL":
                if current_block:
                    blocks.append(current_block)
                break

            if re.match(r"^\d+$", line):
                if current_block:
                    blocks.append(current_block)
                    current_block = []
                current_block.append(line)
            else:
                current_block.append(line)

        if not blocks:
            print(f"⚠️ No rows can be extracted from '{os.path.basename(pdf_path)}' – no blocks found.")
            return []

        # -------------------
        # Extract rows per block
        # -------------------
        for blk in blocks:
            clean_blk = [re.sub(r",", "", l) for l in blk]
            used_lines = set()

            item_code = clean_blk[0] if clean_blk else ""
            used_lines.add(item_code)

            qty, uom = "", ""
            for l in clean_blk:
                if l in used_lines:
                    continue
                m = re.match(r"^(\d+)\s+(\w+)?", l)
                if m:
                    qty = m.group(1)
                    uom = m.group(2) if m.group(2) else ""
                    used_lines.add(l)
                    break

            floats = []
            for l in clean_blk:
                if l in used_lines:
                    continue
                floats += re.findall(r"\d+\.\d+", l)
            price = floats[0] if floats else ""
            total = floats[-1] if floats else ""

            delivery_date = ""
            for l in clean_blk:
                if l in used_lines:
                    continue
                date_match = re.search(r"\d{1,2}/\d{1,2}/\d{4}", l)
                if date_match:
                    delivery_date = date_match.group(0)
                    used_lines.add(l)
                    break

            desc = ""
            usd_idx = None
            for i, l in enumerate(clean_blk):
                if l in used_lines:
                    continue
                if "USD" in l.upper():
                    usd_idx = i
                    used_lines.add(l)
                    break
            if usd_idx is not None and usd_idx + 1 < len(clean_blk):
                desc = " ".join(clean_blk[usd_idx+1:]).strip()
            elif usd_idx is None:
                desc = " ".join(clean_blk[1:]).strip()

            row = {
                "Document Number": document_number,
                "Part/Description": desc,
                "Item_Code": item_code,
                "PO_issue Date": po_issue_date,
                "Delivery Date": delivery_date,
                "SHIP VIA": ship_via,
                "FOB": fob,
                "UoM(optional)": uom,
                "Quantity": qty,
                "Price": price,
                "Total": total,
                "Payment Term": payment_term,
                "BUYER": buyer,
                "REQ#": req_num,
                "REQUISITIONER": requisitioner
            }
            rows.append(row)

    except Exception as e:
        print(f"❌ Error in {pdf_path}: {e}")
    return rows

# -----------------------------
# Main function
# -----------------------------
def main():
    input_dir = Path("input_pdfs")
    output_file = Path("po_extracted.csv")

    if output_file.exists():
        output_file.unlink()

    # Unzip ZIPs & collect PDFs
    pdf_files = []
    for f in input_dir.iterdir():
        if f.suffix.lower() == ".zip":
            with zipfile.ZipFile(f, 'r') as zip_ref:
                zip_ref.extractall(input_dir)
        elif f.suffix.lower() == ".pdf":
            pdf_files.append(str(f))

    for root, _, files_in_dir in os.walk(input_dir):
        for f in files_in_dir:
            if f.lower().endswith(".pdf") and "__MACOSX" not in root:
                full_path = os.path.join(root, f)
                if full_path not in pdf_files:
                    pdf_files.append(full_path)

    all_rows = []
    for pdf_file in pdf_files:
        all_rows.extend(extract_po_info(pdf_file))

    if all_rows:
        df = pd.DataFrame(all_rows)
        df.drop_duplicates(inplace=True)
        df.sort_values(by="Document Number", inplace=True)
        df.to_csv(output_file, index=False)
        print(f"✅ Extraction complete → {output_file}")
    else:
        print("⚠️ No rows extracted.")

if __name__ == "__main__":
    main()
