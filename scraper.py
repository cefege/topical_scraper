import chardet
import trafilatura
import pandas as pd
import re
from io import BytesIO
import xlsxwriter
import openpyxl
import argparse


def move_sheet_to_first(workbook, sheet_name):
    # Get the list of sheet names
    sheet_names = workbook.sheetnames

    # Get the index of the sheet you want to move
    sheet_index = sheet_names.index(sheet_name)

    # Reorder the sheets list to move the desired sheet to the first position
    new_order = [sheet_index] + \
        [i for i in range(len(sheet_names)) if i != sheet_index]

    # Rearrange the sheets in the workbook based on the new order
    workbook._sheets = [workbook._sheets[i] for i in new_order]


def get_correct_encoding(content):
    detected_encoding = chardet.detect(content)
    return detected_encoding["encoding"]


def clean_text(text):
    # Replace non-ASCII characters with their Unicode escape sequences
    cleaned_text = text.encode("ascii", "backslashreplace").decode("utf-8")
    return cleaned_text


def extract_article_headers(xml_content):
    # print(xml_content)
    headers_data = []

    # Extract h1 header if available
    header_tags = re.findall(r'<head rend="(h1)">(.*?)<\/head>', xml_content)
    if header_tags:
        cleaned_header_text = re.sub(r"<[^>]*>", "", header_tags[0][1])
        headers_data.append(
            {"Headings": cleaned_header_text.strip(), "H": "h1"})
    else:
        # Extract title attribute from the main tag if h1 header not found
        title_match = re.search(r'<doc[^>]*title="([^"]+)"', xml_content)
        if title_match:
            title_text = title_match.group(1).strip()
            headers_data.append({"Headings": title_text, "H": "h1"})

    # Process the rest of the header tags
    header_tags = re.findall(r'<head rend="(h\d+)">(.*?)<\/head>', xml_content)
    for header_type, header_text in header_tags:
        if header_type != "h1":
            cleaned_header_text = re.sub(r"<[^>]*>", "", header_text)
            headers_data.append(
                {"Headings": cleaned_header_text.strip(), "H": header_type})

    return headers_data


def clean_headers_dataframe(headers_df):
    # Remove unwanted headers such as "Table of Contents"
    headers_df = headers_df[
        ~headers_df["Headings"].str.lower().str.contains("table of contents")
    ]

    # Drop duplicates
    headers_df.drop_duplicates(subset=["Headings"], inplace=True)

    return headers_df


def url_to_markdown(url):
    print(url)

    html_content = trafilatura.fetch_url(url)
    # Extract and clean the textual content using trafilatura
    xml_content = trafilatura.extract(
        html_content,
        include_formatting=True,
        include_links=True,
        include_tables=True,
        include_images=True,
        output_format="xml",
    )
    return xml_content


def create_excel(urls):
    # Create an in-memory Excel workbook
    excel_output = BytesIO()
    with pd.ExcelWriter(excel_output, engine="xlsxwriter") as writer:
        summary_data = []

        for idx, url in enumerate(urls, start=1):
            try:

                markdown_content = url_to_markdown(url)
                headers_data = extract_article_headers(markdown_content)
                headers_df = pd.DataFrame(headers_data)
                cleaned_headers_df = clean_headers_dataframe(headers_df.copy())

                # Add h1 headers and URLs to the summary dataframe
                h1_header = cleaned_headers_df.loc[
                    cleaned_headers_df["H"] == "h1", "Headings"
                ].iloc[0]
                valid_worksheet_name = re.sub(r'[\/:*?"<>|]', "_", h1_header)

                worksheet_name = valid_worksheet_name[:31]

                summary_data.append({"Title": h1_header, "URL": url})

                cleaned_headers_df.to_excel(
                    writer, sheet_name=worksheet_name, index=False
                )
            except Exception as e:
                summary_data.append(
                    {
                        "Title": "Error processing URL",
                        "URL": url + " (" + str(e) + ")",
                    }
                )
                print(f"Error processing URL: {url}\n{e}")

        # Create the summary DataFrame
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name="summary",
                            index=False, startrow=1)

    # Load the generated workbook and move the summary sheet to first
    wb = openpyxl.load_workbook(excel_output)
    move_sheet_to_first(wb, "summary")

    excel_output = BytesIO()
    wb.save(excel_output)
    return excel_output


def main():
    parser = argparse.ArgumentParser(description="Convert URLs to Excel")
    parser.add_argument(
        "-f", "--file", help="Path to the text file containing URLs", required=True
    )
    args = parser.parse_args()

    with open(args.file, "r", encoding="utf-8") as url_file:
        urls = [url.strip() for url in url_file]
    print(urls)
    excel_output = create_excel(urls)

    print(excel_output)
    # Save the Excel file
    output_filename = "new_output.xlsx"
    with open(output_filename, "wb") as output_file:
        output_file.write(excel_output.getvalue())

    print(f"Excel file '{output_filename}' generated successfully.")


if __name__ == "__main__":
    main()
