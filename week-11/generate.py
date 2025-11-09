import os

folder = "materi"  # folder containing your files
output_file = os.path.join(folder, "index.html")

exclude_dirs = {"venv", "__pycache__"}  # folders to skip

file_links = []

# walk through folder recursively
for root, dirs, files in os.walk(folder):
    # remove excluded directories from traversal
    dirs[:] = [d for d in dirs if d not in exclude_dirs]

    for file in files:
        if file.startswith("."):  # skip hidden files
            continue
        # create relative path for HTML link
        rel_dir = os.path.relpath(root, folder)
        rel_file = os.path.join(rel_dir, file) if rel_dir != "." else file
        file_links.append(rel_file)

file_links.sort()

# generate index.html
with open(output_file, "w", encoding="utf-8") as f:
    f.write("<!DOCTYPE html>\n<html lang='en'>\n<head>\n")
    f.write("<meta charset='UTF-8'>\n<title>Index of Materi</title>\n")
    f.write("</head>\n<body>\n")
    f.write("<h1>Index of Materi</h1>\n<ul>\n")

    for link in file_links:
        f.write(f"<li><a href='/week-11{link}' target='_blank'>{link}</a></li>\n")

    f.write("</ul>\n</body>\n</html>")

print(f"index.html created, listing {len(file_links)} files from '{folder}' (excluding venv and __pycache__)")
