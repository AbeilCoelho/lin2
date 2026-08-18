import hashlib
from pathlib import Path

def sha256_file(filepath, chunk_size=1024 * 1024):
    """Calculate SHA-256 hash of a file."""
    sha256 = hashlib.sha256()

    with filepath.open("rb") as f:
        while chunk := f.read(chunk_size):
            sha256.update(chunk)

    return sha256.hexdigest()


current_dir = Path.cwd()

for file in sorted(current_dir.iterdir()):
    if file.is_file():
        file_hash = sha256_file(file)
        print(f"{file_hash}  {file.name}")