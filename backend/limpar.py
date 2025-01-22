import os
import re
from concurrent.futures import ThreadPoolExecutor
from typing import Iterator

import click

from projetov2.logging import get_logger
from projetov2.utilities.files import find_and_remove_duplicates, remove_files_by_pattern
from projetov2.wrappers import wclick

logger = get_logger(__name__)


def find_and_remove_duplicates(directory: str) -> None:
    # Regular expression to match the pattern "filename (number).extension"
    pattern = re.compile(r"(.+?)\s\(\d+\)(\.\w+)?$")
    files_to_remove = []

    # List all files in the directory
    for root, _, files in os.walk(directory):
        for file in files:
            match = pattern.match(file)
            if match:
                original_filename = match.group(1) + (match.group(2) if match.group(2) else "")
                original_filepath = os.path.join(root, original_filename)
                if os.path.exists(original_filepath):
                    files_to_remove.append(os.path.join(root, file))

    # Remove the duplicate files
    for file_path in files_to_remove:
        os.remove(file_path)
        logger.info(f"Removed duplicate file: {file_path}")


def remove_files_by_pattern(directory: str, pattern: str) -> None:
    """Remove files in the specified directory that match the given glob pattern."""
    # Create the full path pattern
    full_pattern = os.path.join(directory, pattern)

    # Find all files matching the pattern
    files_to_remove = glob.glob(full_pattern)

    # Remove each file
    for file_path in files_to_remove:
        os.remove(file_path)
        logger.info(f"Removed file: {file_path}")


def fast_walk(path: str) -> tuple[str, list[str], list[str]]:
    try:
        with os.scandir(path) as it:
            dirs, files = [], []
            for entry in it:
                if entry.is_dir(follow_symlinks=False):
                    dirs.append(entry.path)  # Store full path
                else:
                    files.append(entry.name)  # Store file name
            return path, dirs, files
    except PermissionError:
        # Skip directories you don't have permissions to access
        return path, [], []


def fast_walk_concurrent(root_dir: str) -> Iterator[tuple[str, list[str], list[str]]]:
    with ThreadPoolExecutor() as executor:
        stack = [root_dir]
        while stack:
            # Submit jobs for directories in the stack
            futures = [executor.submit(fast_walk, d) for d in stack]
            stack = []  # Clear stack to refill with subdirectories
            for future in futures:
                path, dirs, files = future.result()
                yield path, [os.path.basename(d) for d in dirs], files
                stack.extend(dirs)  # Add subdirectories to stack for next level


@wclick.command()
@click.argument("directory", type=wclick.Directory())
def limpar(directory: str) -> None:
    """Remove arquivos desnecessários e realiza processos de limpeza de rotina"""
    for root, dirs, files in fast_walk_concurrent(directory):
        for dirname in dirs:
            dirpath = os.path.join(root, dirname)
            find_and_remove_duplicates(dirpath)
            remove_files_by_pattern(dirpath, "*.tmp")
