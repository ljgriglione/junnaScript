import argparse
from docx import Document
import re
# HAVE TO FIX: be able to handle italics/bold, but change font back..., Also cannot edit things in table
# Define keywords and their corresponding additions
keywords = {
    "signature": "[sig|req|signer1]",
    "name": "[fullname|req|signer1]",
    "date": "[date|req|signer1]",
    "e-signature": "[sig|req|signer1]",
    "e-sig": "[sig|req|signer1]",
    "esig": "[sig|req|signer1]",
    "address": "[text|req|signer1]",
    "phone": "[text|req|signer1]",
    "email": "[text|req|signer1]",
    "initial": "[initial|req|signer1]",
    "title": "[text|req|signer1]",
    "relationship": "[text|req|signer1]",
    "company": "[text|req|signer1]",
    "organization": "[text|req|signer1]",
    "organization name": "[text|req|signer1]",
    "organization title": "[text|req|signer1]",
    "fax": "[text|req|signer1]",
    "home": "[text|req|signer1]",
    "mobile": "[text|req|signer1]",
    "number": "[text|req|signer1]",
    "client": "[text|req|signer1]",
    "Date of birth": "[text|req|signer1]",
    "comments": "[text|req|signer1]",
    "telephone": "[text|req|signer1]"


}

# Set up argument parsing
parser = argparse.ArgumentParser(description="Modify a Word document by adding text next to specific keywords.")
parser.add_argument("input_file", help="Path to the input Word document (.docx)")
args = parser.parse_args()

# Load the Word document
doc = Document(args.input_file)

# Process each paragraph
for para in doc.paragraphs:
    original_text = para.text.strip()  # Preserve original formatting

    # Remove underscores
    modified_text = original_text.replace("_", "")

    # Check if the line is a checkbox option
    if "(check one)" in modified_text.lower():
        # Append [check|noreq|signer1] next to (check one) instead of replacing it
        modified_text = re.sub(r"\(check one\)", "(check one) [check|noreq|signer1]", modified_text, flags=re.IGNORECASE)

    # Apply changes only if the keyword has ":" or is alone
    for word, addition in keywords.items():
        # Match words followed by ":" or empty space (alone in a line)
        pattern = rf"\b({word})([:]|$)"
        modified_text = re.sub(pattern, rf"\1\2 {addition}", modified_text, flags=re.IGNORECASE)

    para.text = modified_text.strip()  # Update paragraph text

# Save changes to the file
doc.save(args.input_file)

print(f"Processing complete. Updated document saved as '{args.input_file}'.")
