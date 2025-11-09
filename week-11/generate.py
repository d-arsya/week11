import os

script_dir = os.path.dirname(os.path.abspath(__file__))

# Path to materi folder relative to the script
folder = os.path.join(script_dir, "materi")
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
    f.write("""<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Index of Materi</title>
    <script src="https://cdn.tailwindcss.com"></script>
  </head>
  <body class="bg-gray-50 text-gray-800 min-h-screen flex flex-col items-center py-10">
    <div class="w-full max-w-2xl bg-white shadow-lg rounded-2xl p-6">
      <h1 class="text-2xl font-bold text-center mb-6 text-indigo-600">📚 Index of Materi</h1>
      <ul class="divide-y divide-gray-200">
""")

    for link in file_links:
        if ".py" not in link:
            f.write(
                f"        <li class='py-2 hover:bg-indigo-50 px-3 rounded-md transition'>"
                f"<a href='/week-11/materi/{link}' target='_blank' class='text-indigo-600 hover:text-indigo-800 font-medium'>{link}</a>"
                f"</li>\n"
            )

    f.write("""      </ul>
      <p class="text-center text-sm text-gray-400 mt-6">Generated automatically</p>
    </div>
  </body>
</html>""")


print(f"index.html created, listing {len(file_links)} files from '{folder}' (excluding venv and __pycache__)")
