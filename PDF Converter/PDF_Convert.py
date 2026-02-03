import os, comtypes.client

FOLDER = r"D:\PARTH\UNFINISHED PROJECTS\Dataset Raw"


def convert_docs(word_app, doc_files):
    for file_path in doc_files:
        name, _ = os.path.splitext(file_path)
        pdf_path = name + ".pdf"
        try:
            doc = word_app.Documents.Open(file_path)
            doc.SaveAs(pdf_path, FileFormat=17)
            doc.Close()
            os.remove(file_path)
            print(f"Converted Word: {file_path}")
        except Exception as e:
            print(f"Error converting Word {file_path}: {e}")


def convert_ppts(ppt_app, ppt_files):
    for file_path in ppt_files:
        name, _ = os.path.splitext(file_path)
        pdf_path = name + ".pdf"
        try:
            pres = ppt_app.Presentations.Open(file_path)
            pres.SaveAs(pdf_path, 32)
            pres.Close()
            os.remove(file_path)
            print(f"Converted PowerPoint: {file_path}")
        except Exception as e:
            print(f"Error converting PowerPoint {file_path}: {e}")


def main():
    doc_files = []
    ppt_files = []
    for root, _, files in os.walk(FOLDER):
        for file in files:
            if file.lower().endswith((".doc", ".docx")):
                doc_files.append(os.path.join(root, file))
            elif file.lower().endswith((".ppt", ".pptx")):
                ppt_files.append(os.path.join(root, file))
            elif not file.lower().endswith(".pdf"):
                print(f"Skipping unsupported file: {os.path.join(root,file)}")
    word_app = comtypes.client.CreateObject("Word.Application")
    word_app.Visible = False
    try:
        convert_docs(word_app, doc_files)
    finally:
        word_app.Quit()
    ppt_app = comtypes.client.CreateObject("Powerpoint.Application")
    ppt_app.Visible = True
    ppt_app.WindowState = 2
    try:
        convert_ppts(ppt_app, ppt_files)
    finally:
        ppt_app.Quit()


if __name__ == "__main__":
    main()
