import os
import pdfplumber
import pandas as pd
import camelot
import streamlit as st
from typing import List, Optional, Union


def clean_dataframe(df: pd.DataFrame, expected_headers: List[str]) -> pd.DataFrame:
    """Clean the dataframe by removing duplicate headers and empty rows."""
    if df is None or df.empty:
        return pd.DataFrame(columns=expected_headers)

    # Drop rows where the first column matches the header, if duplicated
    df_cleaned = df[~(df.iloc[:, 0] == expected_headers[0])]
    return df_cleaned.dropna(how='all').reset_index(drop=True)


def extract_page_data_with_pdfplumber(
        page: pdfplumber.page.Page,
        column_positions: List[float],
        expected_headers: List[str]
) -> Optional[pd.DataFrame]:
    """Extract data from a PDF page using pdfplumber."""
    try:
        words = page.extract_words()
        rows = {}

        for word in words:
            y0 = word['top']
            text = word['text']
            x0 = word['x0']
            row_key = round(y0, -1)

            if row_key not in rows:
                rows[row_key] = [[] for _ in range(len(expected_headers))]

            for idx, (start, end) in enumerate(zip(column_positions, column_positions[1:] + [None])):
                if (x0 >= start) and (end is None or x0 < end):
                    rows[row_key][idx].append(text)
                    break

        data_rows = [
            [" ".join(cell).strip() for cell in rows[row_key]]
            for row_key in sorted(rows.keys())
        ]

        if not data_rows:
            return None

        df = pd.DataFrame(data_rows, columns=expected_headers)
        return clean_dataframe(df, expected_headers)

    except Exception as e:
        st.error(f"Error with pdfplumber extraction: {str(e)}")
        return None


def process_pdf(
        pdf_path: str,
        output_path: str,
        pages: Union[str, List[int]] = 'all',
        start_page: Optional[int] = None,
        end_page: Optional[int] = None,
        password: Optional[str] = None
) -> Optional[pd.DataFrame]:
    """Process PDF and extract data into a DataFrame."""

    expected_headers = [
        'Receipt No.',
        'Completion Time',
        'Details',
        'Transaction Status',
        'Paid In',
        'Withdrawn',
        'Balance'
    ]
    column_positions = [37.5, 85, 194.899, 350, 418.4, 465.2, 521.34]

    all_data = []

    try:
        with pdfplumber.open(pdf_path, password=password) as pdf:
            total_pages = len(pdf.pages)
            st.info(f"Total pages in PDF: {total_pages}")

            if pages == 'all':
                if start_page is not None and end_page is not None:
                    pages_to_process = range(
                        max(start_page - 1, 0),
                        min(end_page, total_pages)
                    )
                else:
                    pages_to_process = range(total_pages)
            else:
                pages_to_process = [p - 1 for p in pages if 0 < p <= total_pages]

            st.info(f"Processing pages: {[p + 1 for p in pages_to_process]}")

            for page_num in pages_to_process:
                try:
                    page = pdf.pages[page_num]

                    st.info(f"Attempting pdfplumber extraction on page {page_num + 1}")
                    df_pdfplumber = extract_page_data_with_pdfplumber(
                        page,
                        column_positions,
                        expected_headers
                    )

                    if df_pdfplumber is not None and not df_pdfplumber.empty:
                        all_data.append(df_pdfplumber)
                        st.success(f"Extracted data from page {page_num + 1} with pdfplumber")
                    else:
                        st.info(f"Attempting Camelot extraction on page {page_num + 1}")
                        tables = camelot.read_pdf(pdf_path, pages=str(page_num + 1), password=password)

                        for table in tables:
                            df_camelot = table.df
                            if len(df_camelot.columns) == len(expected_headers):
                                df_camelot.columns = expected_headers
                                df_camelot = clean_dataframe(df_camelot, expected_headers)
                                if df_camelot is not None and not df_camelot.empty:
                                    all_data.append(df_camelot)
                                    st.success(f"Extracted data from page {page_num + 1} with Camelot")

                except Exception as e:
                    st.error(f"Error processing page {page_num + 1}: {str(e)}")
                    continue

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            final_df = clean_dataframe(final_df, expected_headers)
            final_df.to_excel(output_path, index=False)
            return final_df

        st.warning("No data was successfully extracted from the PDF.")
        return None

    except Exception as e:
        st.error(f"An error occurred while processing the PDF: {str(e)}")
        return None


def main():
    """Main Streamlit application."""
    st.title("PDF to Excel Converter")

    uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")
    password = st.text_input("Enter PDF password (if applicable)", type="password")

    if uploaded_file is None:
        return

    st.success("PDF file uploaded successfully.")

    with open("temp.pdf", "wb") as f:
        f.write(uploaded_file.read())

    process_all_pages = st.checkbox("Process all pages", value=True)
    pages, start_page, end_page = 'all', None, None

    if not process_all_pages:
        page_option = st.radio(
            "Choose page selection method:",
            ("Range", "Specific Pages")
        )

        if page_option == "Range":
            start_page = st.number_input("Start page", min_value=1, value=1)
            end_page = st.number_input("End page", min_value=start_page, value=start_page)
        else:
            specific_pages = st.text_input(
                "Enter specific pages (comma-separated, e.g., 1,3,5)"
            )
            if specific_pages:
                try:
                    pages = [int(p.strip()) for p in specific_pages.split(",")]
                except ValueError:
                    st.error("Please enter valid page numbers.")
                    return

    if not st.button("Convert to Excel"):
        return

    output_path = "converted_data.xlsx"

    if process_all_pages:
        df = process_pdf("temp.pdf", output_path, password=password)
    elif start_page and end_page:
        df = process_pdf("temp.pdf", output_path, start_page=start_page, end_page=end_page, password=password)
    else:
        df = process_pdf("temp.pdf", output_path, pages=pages, password=password)

    if df is None:
        st.error("Failed to process PDF.")
        return

    df = pd.read_excel(output_path)
    df_filtered = df[df['Balance'].notna()]
    excel_output_path = "clean.xlsx"
    df_filtered.to_excel(excel_output_path, index=False)

    st.write("Preview of extracted data:")
    st.write(df_filtered.head())

    with open(excel_output_path, "rb") as excel_file:
        st.download_button(
            label="Download Excel file",
            data=excel_file,
            file_name="clean.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"  )

    for file in ["temp.pdf", "converted_data.xlsx", "clean.xlsx"]:
        if os.path.exists(file):
            os.remove(file)


if __name__ == "__main__":
    main()
