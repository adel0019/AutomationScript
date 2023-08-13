from os import scandir, rename
from os.path import splitext, exists, join
from shutil import move
from time import sleep
import zipfile
import os
import logging

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

source_dir = "/Users/Mohamed/Onedrive"
dest_dir_image = "/Users/Mohamed/OneDrive/Pictures/Downloaded Images"
dest_dir_documents = "/Users/Mohamed/OneDrive/Documents/Downloaded Docs"
dest_dir_installations = "/Users/Mohamed/Onedrive/Installations"
dest_dir_zipfiles = "/Users/Mohamed/Onedrive/Extracted Zip Files"

#  supported image types
image_extensions = [".jpg", ".jpeg", ".jpe", ".jif", ".jfif", ".jfi", ".png", ".gif", ".webp", ".tiff", ".tif", ".psd",
                    ".raw", ".arw", ".cr2", ".nrw",
                    ".k25", ".bmp", ".dib", ".heif", ".heic", ".ind", ".indd", ".indt", ".jp2", ".j2k", ".jpf", ".jpf",
                    ".jpx", ".jpm", ".mj2", ".svg", ".svgz", ".ai", ".eps", ".ico"]

#  supported Document types
document_extensions = [".doc", ".docx", ".odt",
                       ".pdf", ".xls", ".xlsx", ".ppt", ".pptx"]


def make_unique(dest, name):
    filename, extension = splitext(name)
    counter = 1
    while exists(f"{dest}/{name}"):
        name = f"{filename}({str(counter)}){extension}"
        counter += 1

    return name


def move_file(dest, entry, name):
    if exists(f"{dest}/{name}"):
        unique_name = make_unique(dest, name)
        oldName = join(dest, name)
        newName = join(dest, unique_name)
        rename(oldName, newName)
    move(entry, dest)


class MoverHandler(FileSystemEventHandler):
    def on_modified(self, event):
        with scandir(source_dir) as entries:
            for entry in entries:
                name = entry.name
                self.check_image_files(entry, name)
                self.check_document_files(entry, name)
                self.check_installation_files(entry, name)
                self.check_zip_files(entry, name)

    def on_created(self, event):
        # Add check for zip files
        if not event.is_directory:
            entry = event.src_path
            name = os.path.basename(entry)
            self.check_zip_files(entry, name)

    def check_installation_files(self, entry, name):
        if name.endswith('.exe') or name.endswith('.msi'):
            move_file(dest_dir_installations, entry, name)
            logging.info(f"Moved image file: {name}")

    def check_image_files(self, entry, name):
        for image_extension in image_extensions:
            if name.endswith(image_extension) or name.endswith(image_extension.upper()):
                move_file(dest_dir_image, entry, name)
                logging.info(f"Moved image file: {name}")

    def check_document_files(self, entry, name):
        for documents_extension in document_extensions:
            if name.endswith(documents_extension) or name.endswith(documents_extension.upper()):
                move_file(dest_dir_documents, entry, name)
                logging.info(f"Moved document file: {name}")

    def check_zip_files(self, entry, name):
        print(f"Checking zip file: {name}")
        if name.lower().endswith('.zip'):
            try:
                print(f"Processing zip file: {name}")
                with zipfile.ZipFile(entry, 'r') as zip_ref:
                    zip_ref.extractall(dest_dir_zipfiles)
                os.remove(entry)
                logging.info(f"Extracted and deleted zip file: {name}")
            except Exception as e:
                logging.error(f"Error while processing zip file {name}: {e}")


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s - %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S')
    path = source_dir
    event_handler = MoverHandler()
    observer = Observer()
    observer.schedule(event_handler, path, recursive=True)
    observer.start()
    try:
        while True:
            sleep(10)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
