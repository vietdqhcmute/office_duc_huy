def update_run_text(run, ws, r_idx, dict_cl_name):
    if run.text in dict_cl_name.keys():
        content = icell(ws, r_idx, dict_cl_name[run.text])
        run.text = content if content != "None" else ""

def process_paragraphs(paragraphs, ws, r_idx, dict_cl_name):
    for para in paragraphs:
        for run in para.runs:
            update_run_text(run, ws, r_idx, dict_cl_name)

def process_tables(tables, ws, r_idx, dict_cl_name):
    for tab in tables:
        for row in tab.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs, ws, r_idx, dict_cl_name)

def main():
    # ... existing code ...

    for r_idx in range(2, ws.max_row + 1):
        out_name = lcell(ws, r_idx, name_column_letter) + ".docx"
        shutil.copy(doc_sample_path, join(des_folder, out_name))

        doc = docx.Document(join(des_folder, out_name))
        process_paragraphs(doc.paragraphs, ws, r_idx, dict_cl_name)
        process_tables(doc.tables, ws, r_idx, dict_cl_name)

    # ... existing code ...

if __name__ == "__main__":
    main()
