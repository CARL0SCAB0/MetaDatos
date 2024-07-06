import os
import sys
import docx
import openpyxl
import PyPDF2

def extract_metadata_docx(file_path):
    doc = docx.Document(file_path)
    metadata = doc.core_properties
    return {
        "Author": metadata.author,
        "Title": metadata.title,
        "Subject": metadata.subject,
        "Keywords": metadata.keywords,
        "Last Modified By": metadata.last_modified_by,
        "Created": metadata.created,
        "Modified": metadata.modified,
    }

def extract_metadata_xlsx(file_path):
    workbook = openpyxl.load_workbook(file_path)
    properties = workbook.properties
    return {
        "Title": properties.title,
        "Subject": properties.subject,
        "Creator": properties.creator,
        "Keywords": properties.keywords,
        "Last Modified By": properties.lastModifiedBy,
        "Created": properties.created,
        "Modified": properties.modified,
    }

def extract_metadata_pdf(file_path):
    with open(file_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        metadata = reader.metadata
        return {
            "Title": metadata.title,
            "Author": metadata.author,
            "Subject": metadata.subject,
            "Creator": metadata.creator,
            "Producer": metadata.producer,
            "Created": metadata.creation_date,
            "Modified": metadata.modification_date,
        }

def extract_metadata(file_path):
    if file_path.endswith('.docx'):
        return extract_metadata_docx(file_path)
    elif file_path.endswith('.xlsx'):
        return extract_metadata_xlsx(file_path)
    elif file_path.endswith('.pdf'):
        return extract_metadata_pdf(file_path)
    else:
        return None

def main(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(('.docx', '.xlsx', '.pdf')):
                file_path = os.path.join(root, file)
                metadata = extract_metadata(file_path)
                if metadata:
                    print(f"Metadata for {file_path}:")
                    for key, value in metadata.items():
                        print(f"  {key}: {value}")
                    print("\n")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python extract_metadata.py <directory>")
        sys.exit(1)
    
    directory = sys.argv[1]
    if not os.path.isdir(directory):
        print(f"The directory {directory} does not exist.")
        sys.exit(1)
    
    main(directory)
