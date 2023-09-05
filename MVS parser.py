import os
import xml.etree.ElementTree as ET
import pandas as pd
import PyPDF2

def get_page_count(pdf_path):
    """Get the page count of a PDF."""
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        return len(reader.pages)

def parse_xml(xml_path):
    """Parse the XML and extract required details."""
    tree = ET.parse(xml_path)
    root = tree.getroot()
    
    # Extracting the required details
    volume_count = root.find(".//BookMultiVolumeCount").text
    split_after_chapter = root.find(".//BookMultiVolumeSplitAfterChapter").text
    total_pages = [elem.text for elem in root.findall(".//CompoundObjectTotalNumberOfPages")]
    
    return volume_count, split_after_chapter, total_pages

def main():
    # Define the main directory containing the book folders
    main_dir = "D:\Springer MVS\To Check"  # <- Replace this with the path to your main directory
    
    data = []
    
    for book in os.listdir(main_dir):
        book_path = os.path.join(main_dir, book)
        
        # Get number of volumes and their page counts
        pdf_dir = os.path.join(book_path, "BodyRef", "PDF")
        pdf_files = [f for f in os.listdir(pdf_dir) if f.endswith('.pdf')]
        if pdf_files:
            volumes = [1]
        else:
            volumes = [v for v in os.listdir(pdf_dir) if os.path.isdir(os.path.join(pdf_dir, v))]


        # If there are PDF files directly in the pdf_dir, treat it as a single volume
        if pdf_files:
            page_counts = [get_page_count(os.path.join(pdf_dir, pdf_files[0]))]
        else:
            page_counts = []
            for v in volumes:
                # Get the first PDF file inside the volume directory (assuming there's only one PDF per volume)
                pdf_file = next(f for f in os.listdir(os.path.join(pdf_dir, v)) if f.endswith('.pdf'))
                page_counts.append(get_page_count(os.path.join(pdf_dir, v, pdf_file)))
        
        # Parse the XML file
        xml_file = [f for f in os.listdir(book_path) if f.endswith(".xml")][0]
        volume_count, split_after_chapter, total_pages = parse_xml(os.path.join(book_path, xml_file))
        
        data.append([book, len(volumes), page_counts, volume_count, split_after_chapter, total_pages])
        total_page_count = sum(page_counts)
        total_pages_xml = sum([int(p) for p in total_pages])
        data[-1].extend([total_page_count, total_pages_xml])
        
    # Create a DataFrame and save to Excel
    df = pd.DataFrame(data, columns=["ISBN", "Number of Volumes", "Page Counts", "Volume Count from XML", "Split After Chapter", "Total Pages from XML", "Total Page Count", "Sum of Pages from XML"])
    df.to_excel("book_details.xlsx", index=False)

if __name__ == "__main__":
    main()