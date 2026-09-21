import re
from pathlib import Path
from docx import Document

# Replace with your ACTUAL root constant/import — see note at the end
DOCUMENTS_ROOT = Path(__file__).resolve().parent.parent / "documents"


def sanitize_name(name: str) -> str:
    name = name.strip()
    name = re.sub(r'[\\/:*?"<>|]', "", name)
    name = name.replace("..", "")
    return name


def create_new_repo(repo_name: str):
    """Creates a new top-level folder directly under DOCUMENTS_ROOT."""
    clean_name = sanitize_name(repo_name)
    if not clean_name:
        return False, "Invalid repo name.", None

    target_path = DOCUMENTS_ROOT / clean_name
    if target_path.exists():
        return False, f"A repo named '{clean_name}' already exists.", None

    target_path.mkdir(parents=True)
    return True, "Repo created.", clean_name


def create_new_folder(directory_path, folder_name: str):
    """Creates a new subfolder inside an existing repo/folder."""
    directory_path = Path(directory_path)
    if not directory_path.exists() or not directory_path.is_dir():
        return False, "Target folder does not exist.", None

    clean_name = sanitize_name(folder_name)
    if not clean_name:
        return False, "Invalid folder name.", None

    target_path = directory_path / clean_name
    if target_path.exists():
        return False, f"A folder named '{clean_name}' already exists.", None

    target_path.mkdir(parents=True)
    return True, "Folder created.", clean_name


def create_new_docx(directory_path, file_name: str):
    """Creates a new blank .docx inside an existing repo/folder."""
    directory_path = Path(directory_path)
    if not directory_path.exists() or not directory_path.is_dir():
        return False, "Target folder does not exist.", None

    clean_name = sanitize_name(file_name)
    if not clean_name:
        return False, "Invalid file name.", None
    if not clean_name.lower().endswith(".docx"):
        clean_name += ".docx"

    target_path = directory_path / clean_name
    if target_path.exists():
        return False, f"A file named '{clean_name}' already exists.", None

    doc = Document()
    doc.add_paragraph("")
    doc.save(target_path)

    return True, "Document created.", clean_name


def resolve_folder_path(subpath: str) -> Path:
    """Resolves subpath to a real directory under DOCUMENTS_ROOT, blocking traversal."""
    target = (DOCUMENTS_ROOT / subpath).resolve()
    root_resolved = DOCUMENTS_ROOT.resolve()

    if root_resolved != target and root_resolved not in target.parents:
        raise FileNotFoundError(f"{subpath} is outside the documents root")

    if not target.exists() or not target.is_dir():
        raise FileNotFoundError(f"{target} not exist")

    return target