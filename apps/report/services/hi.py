from docx import Document


def get_column_indices(table, required_headers):
    header_row = table.rows[0]
    headers = [cell.text.strip() for cell in header_row.cells]
    return [headers.index(header) for header in required_headers if header in headers]


def get_indices(doc, test_type):
    indices = {}
    search_terms = {
        "WAIS": ["WAIS-V Index", "WAIS-IV Subtest"],
        "Brown": ["Brown Score Summary"]
    }
    terms_to_search = search_terms.get(test_type, [])

    for index, table in enumerate(doc.tables):
        for row in table.rows:
            for cell in row.cells:
                if any(term in cell.text for term in terms_to_search):
                    if "WAIS-V Index" in cell.text:
                        indices["WAIS-V Index"] = {
                            "table_index": index,
                            "columns": get_column_indices(table, ["Composite Score", "Percentile", "Description"])
                        }
                    elif "WAIS-IV Subtest" in cell.text:
                        indices["WAIS-IV Subtest"] = {
                            "table_index": index,
                            "columns": get_column_indices(table, ["Scaled Score", "Percentile"])
                        }
                    elif "Brown Score Summary" in cell.text:
                        indices["Brown Score Summary"] = {
                            "table_index": index,
                            "columns": get_column_indices(table, ["T Score", "Percentile"])
                        }
    return indices


# doc_path = '/home/ubuntu/flask/files/ADHD.docx'
# doc = Document(doc_path)
# print(get_indices(doc, "WAIS"))
