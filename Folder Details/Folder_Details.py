import os, csv, statistics
from pypdf import PdfReader
import win32com.client


def get_pdf_page_count(path):
    try:
        reader = PdfReader(path)
        return len(reader.pages)
    except Exception as e:
        print(f"Warning: could not read PDF {path}: {e}")
        return 0


def get_word_page_count(word_app, path):
    try:
        doc = word_app.Documents.Open(path, ReadOnly=True)
        doc.Repaginate()
        pages = doc.ComputeStatistics(2)
        doc.Close(SaveChanges=False)
        return int(pages)
    except Exception as e:
        print(f"Warning: could not read Word file {path}: {e}")
        return 0


def get_ppt_slide_count(ppt_app, path):
    try:
        pres = ppt_app.Presentations.Open(path, WithWindow=False)
        slides = len(pres.Slides)
        pres.Close()
        return int(slides)
    except Exception as e:
        print(f"Warning: could not read PowerPoint file {path}: {e}")
        return 0


def get_folder_details_to_csv(parent_folder):
    output_file = os.path.join(parent_folder, "Folder_Details.csv")
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True
    try:
        with open(output_file, mode="w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(
                ["Subfolder Name", "Total Pages/Slides", "File Count", "Median"]
            )
            for subfolder in os.listdir(parent_folder):
                full_path = os.path.join(parent_folder, subfolder)
                if not os.path.isdir(full_path):
                    continue
                page_counts = []
                for dirpath, _, filenames in os.walk(full_path):
                    for f in filenames:
                        fp = os.path.join(dirpath, f)
                        if not os.path.isfile(fp):
                            continue
                        ext = os.path.splitext(f)[1].lower()
                        if ext == ".pdf":
                            count = get_pdf_page_count(fp)
                        elif ext in (".doc", ".docx"):
                            count = get_word_page_count(word_app, fp)
                        elif ext in (".ppt", ".pptx"):
                            count = get_ppt_slide_count(ppt_app, fp)
                        else:
                            continue
                        if count > 0:
                            page_counts.append(count)
                if page_counts:
                    total_pages = sum(page_counts)
                    file_count = len(page_counts)
                    median_val = statistics.median(page_counts)
                else:
                    total_pages = 0
                    file_count = 0
                    median_val = 0
                writer.writerow([subfolder, total_pages, file_count, median_val])
        print(f"Data exported to: {output_file}")
    finally:
        try:
            word_app.Quit()
        except Exception:
            pass
        try:
            ppt_app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    get_folder_details_to_csv(
        r"D:\PARTH\MY STUFF\VIT\SEM 6\BCSE301L Software Engineering"
    )
