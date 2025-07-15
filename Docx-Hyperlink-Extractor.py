from datetime import datetime
from docx import Document
import csv
import re
import os


def extract_links_with_line_numbers(docx_path):
    document = Document(docx_path)
    url_regex = re.compile(r'https?://\S+')

    results = []
    processed = []
    for idx, para in enumerate(document.paragraphs, start=1):
        text = para.text.strip()
        if not text:
            continue

        # Extract embedded hyperlinks
        for hyperlink in para._element.xpath('.//w:hyperlink'):
            r_id = hyperlink.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if r_id:
                link = document.part.rels[r_id].target_ref
                display_text = ''.join(node.text for node in hyperlink.xpath('.//w:t') if node.text)
                # print(f"[Embedded Hyperlink] - Line#: {idx} - Text: {display_text} Link: {link}")
                processed.append(idx)
                results.append(tuple(["Embedded Hyperlink",idx,display_text,link,]))

        # Extract plain URLs
        urls_in_text = url_regex.findall(text)
        for url in urls_in_text:
            if idx in processed:
                pass
            else:
                # print(f"[Plain URL] - Line#: {idx} - Link: {url}")
                processed.append(idx)
                results.append(tuple(["Plain URL",idx,"",url,]))
        
        else:
            if idx in processed:
                pass
            else:
                # print(f"[Text Only] - Line#: {idx} - Text: {text}")
                results.append(tuple(["Text Only",idx,text,""]))

    return results

def csv_file_maker(results):
    user = user = os.getlogin()
    current_date = datetime.now()
    formatted_date = current_date.strftime("%Y-%m-%dT%H-%M-%S")
    filename = f"C:\\Users\\{user}\\Downloads\\extracted_links__{formatted_date}.csv"

    with open(filename, 'w', newline='') as csvfile:
        # Create a csv.writer object
        csv_writer = csv.writer(csvfile)

        # Optionally, write a header row
        header = ["Category", "Line Number", "Text", "Link"]
        csv_writer.writerow(header)

        # Write the data rows
        csv_writer.writerows(results)

if __name__ == "__main__":
    
    file_picked = input("Enter in the absolute filepath: ")
    if file_picked[0] == '"':
        file_picked = file_picked[1:-1]
    results = extract_links_with_line_numbers(file_picked)
    csv_file_maker(results)


"""
References:
- https://chatgpt.com/[REDACTED]
- https://www.google.com/search?q=python+get+windows+user+logged+in
"""
