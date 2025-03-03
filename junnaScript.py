import argparse
from docx import Document
import re

# Define keywords and their additions
keywords = {
    "signature": " SIGADD ",
    "name": " NAMEADD ",
    "date": " DATEADD ",
    "esignature": " ESIGADD ",
    "e-signature": " ESIGADD "
}

# Set up argument parsing
parser = argparse.ArgumentParser(description="Modify a Word document by adding text next to specific keywords.")
parser.add_argument("input_file", help="Path to the input Word document (.docx)")
args = parser.parse_args()

# Load the Word document
doc = Document(args.input_file)

# Process each paragraph
for para in doc.paragraphs:
    modified_text = para.text  # Preserve original text

    # Remove all underscores from the paragraph
    modified_text = modified_text.replace("_", "")

    # Add text after keywords while keeping punctuation
    for word, addition in keywords.items():
        # Updated regex: Match word and handle any trailing punctuation + spaces
        pattern = rf"\b({word})([:;,.\-]*)(\s*)"
        modified_text = re.sub(pattern, rf"\1\2\3{addition}", modified_text, flags=re.IGNORECASE)

    para.text = modified_text  # Update paragraph text

# Save changes to the same file
doc.save(args.input_file)

print(f"Processing complete. '{args.input_file}' has been updated.")
