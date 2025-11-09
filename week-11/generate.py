import os

folder = "materi"  # folder containing your files
output_file = "materi/index.html"

files = os.listdir(folder)
files.sort()  # optional, sort alphabetically

with open(output_file, "w", encoding="utf-8") as f:
    f.write("<!DOCTYPE html>\n<html lang='en'>\n<head>\n")
    f.write("<meta charset='UTF-8'>\n<title>Index of Materi</title>\n")
    f.write("</head>\n<body>\n")
    f.write("<h1>Index of Materi</h1>\n<ul>\n")

    for file in files:
        # skip hidden files
        if file.startswith("."):
            continue
        file_path = os.path.join(folder, file)
        f.write(f"<li><a href='/week-11/{file_path}' target='_blank'>{file}</a></li>\n")

    f.write("</ul>\n</body>\n</html>")

print(f"index.html created, listing {len(files)} files from '{folder}'")
