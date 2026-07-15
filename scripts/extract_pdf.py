"""Extract text and tables from a procurement announcement PDF.

Usage:
    python extract_pdf.py --input <pdf_path> [--output <output_path>]
"""
import pdfplumber
import argparse
import sys

# Force UTF-8
sys.stdout.reconfigure(encoding='utf-8')


def extract_pdf(pdf_path, output_path):
    """Extract text and tables from a PDF file."""
    with open(output_path, 'w', encoding='utf-8') as f:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    f.write(f'=== Page {i+1} ===\n')
                    f.write(text + '\n\n')

                tables = page.extract_tables()
                for j, table in enumerate(tables):
                    f.write(f'--- Table {j+1} on Page {i+1} ---\n')
                    for row in table:
                        f.write(str(row) + '\n')
                    f.write('\n')

    print(f'Output written to {output_path}')


def main():
    parser = argparse.ArgumentParser(
        description='Extract text and tables from a procurement PDF'
    )
    parser.add_argument('--input', '-i', required=True, help='Path to the input PDF file')
    parser.add_argument('--output', '-o', default='pdf_output.txt',
                        help='Path for the output text file (default: pdf_output.txt)')
    args = parser.parse_args()

    extract_pdf(args.input, args.output)


if __name__ == '__main__':
    main()
