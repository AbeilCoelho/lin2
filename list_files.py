import os

# Extensions to include
EXTENSIONS = [".py", ".html"]

# Current directory
DIRECTORY = "."


def print_files():
    for root, _, files in os.walk(DIRECTORY):
        for filename in sorted(files):

            if not any(filename.lower().endswith(ext) for ext in EXTENSIONS):
                continue

            filepath = os.path.abspath(os.path.join(root, filename))

            print("=" * 100)
            print(f"FILE: {filepath}")
            print("=" * 100)

            try:
                with open(
                    filepath,
                    "r",
                    encoding="utf-8",
                    errors="replace"
                ) as f:
                    print(f.read())

            except Exception as e:
                print(f"[ERROR READING FILE] {e}")

            print()


if __name__ == "__main__":
    print_files()