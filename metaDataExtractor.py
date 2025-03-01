import os
import sys
import json
import exifread
import PyPDF2
import mutagen  # Library for extracting metadata from audio files
import docx  # Library for extracting metadata from Word documents

def extract_image_metadata(image_path):
    """
    Extract metadata from an image file using ExifRead.
    :param image_path: Path to the image file
    :return: Dictionary containing extracted metadata
    """
    metadata = {}
    try:
        with open(image_path, 'rb') as img_file:
            tags = exifread.process_file(img_file)
            metadata = {tag: str(value) for tag, value in tags.items()}
    except Exception as e:
        print(f"Error extracting metadata from image {image_path}: {e}")
    return metadata

def extract_pdf_metadata(pdf_path):
    """
    Extract metadata from a PDF file using PyPDF2.
    :param pdf_path: Path to the PDF file
    :return: Dictionary containing extracted metadata
    """
    metadata = {}
    try:
        with open(pdf_path, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            info = reader.metadata
            if info:
                metadata = {key[1:]: str(value) for key, value in info.items()}  # Remove leading '/'
    except Exception as e:
        print(f"Error extracting metadata from PDF {pdf_path}: {e}")
    return metadata

def extract_audio_metadata(audio_path):
    """
    Extract metadata from an audio file using Mutagen.
    :param audio_path: Path to the audio file
    :return: Dictionary containing extracted metadata
    """
    metadata = {}
    try:
        audio = mutagen.File(audio_path)
        if audio:
            metadata = {key: str(value) for key, value in audio.tags.items()}
    except Exception as e:
        print(f"Error extracting metadata from audio file {audio_path}: {e}")
    return metadata

def extract_doc_metadata(doc_path):
    """
    Extract metadata from a Word document using python-docx.
    :param doc_path: Path to the Word document
    :return: Dictionary containing extracted metadata
    """
    metadata = {}
    try:
        doc = docx.Document(doc_path)
        metadata = {prop: str(getattr(doc.core_properties, prop)) for prop in dir(doc.core_properties) if not prop.startswith('_')}
    except Exception as e:
        print(f"Error extracting metadata from Word document {doc_path}: {e}")
    return metadata

def save_metadata_html(metadata, filename):
    """
    Save extracted metadata as an HTML table for better readability in the script directory.
    :param metadata: Dictionary containing extracted metadata
    :param filename: Name of the original file
    """
    output_path = os.path.join(os.getcwd(), os.path.basename(filename) + "_metadata.html")
    with open(output_path, 'w') as f:
        f.write("<html><head><title>Metadata</title></head><body>")
        f.write("<h2>Extracted Metadata</h2>")
        f.write("<table border='1' style='border-collapse: collapse; width: 100%;'>")
        f.write("<tr><th>Key</th><th>Value</th></tr>")
        for key, value in metadata.items():
            f.write(f"<tr><td>{key}</td><td>{value}</td></tr>")
        f.write("</table></body></html>")
    print(f"Metadata saved to {output_path}")

def save_metadata_json(metadata, filename):
    """
    Save extracted metadata as a JSON file in the script directory.
    :param metadata: Dictionary containing extracted metadata
    :param filename: Name of the original file
    """
    output_path = os.path.join(os.getcwd(), os.path.basename(filename) + "_metadata.json")
    with open(output_path, 'w') as f:
        json.dump(metadata, f, indent=4)
    print(f"Metadata saved to {output_path}")

def analyze_file(file_path):
    """
    Determine the file type and extract metadata accordingly.
    :param file_path: Path to the file to be analyzed
    """
    if not os.path.isfile(file_path):
        print("Invalid file path. Please provide a valid file.")
        return
    
    file_extension = os.path.splitext(file_path)[1].lower()
    
    if file_extension in ['.jpg', '.jpeg', '.png', '.tiff']:
        metadata = extract_image_metadata(file_path)
    elif file_extension == '.pdf':
        metadata = extract_pdf_metadata(file_path)
    elif file_extension in ['.mp3', '.flac', '.wav', '.m4a']:
        metadata = extract_audio_metadata(file_path)
    elif file_extension in ['.docx']:
        metadata = extract_doc_metadata(file_path)
    else:
        print("Unsupported file format. Only images, PDFs, audio files, and Word documents are supported.")
        return
    
    if metadata:
        print("\nMetadata Extracted:")
        for key, value in metadata.items():
            print(f"{key}: {value}")
        
        save_metadata_html(metadata, file_path)
        
        save_json = input("Do you want to save the metadata as a JSON file? (yes/no, press Enter for yes): ").strip().lower()
        if save_json == '' or save_json == 'yes':
            save_metadata_json(metadata, file_path)
    else:
        print("No metadata found in the file.")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = input("Enter the path of the file to analyze: ").strip().strip('"')
    
    analyze_file(file_path)
