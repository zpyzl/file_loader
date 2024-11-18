import os
import sys
from pathlib import Path

from flask import Flask, request, jsonify

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))))))

from chatchat.server.file_rag.document_loaders import RapidOCRDocLoader, RapidOCRPDFLoader
from chatchat.server.file_rag.document_loaders.ofd_loader import read_ofd
from log_utils import setup_log

logger = setup_log('file_loader_service.log',True)

app = Flask(__name__)

DIRPATH = r"D:\test_rag_doc"


@app.route('/load_file', methods=['GET', 'POST'])
def load_file():
    directory = request.json.get('dir_path')
    documents = []
    pdfs, docs = [], []
    loader = None
    converted_pdf_path = None
    for file in Path(directory).rglob('*'):
        if file.is_file():
            try:
                logger.info(f"Processing {file}")
                ext = file.suffix
                if ext in ['.doc', '.docx']:
                    loader = RapidOCRDocLoader(file_path=file)
                elif ext == '.pdf':
                    loader = RapidOCRPDFLoader(file_path=file)
                elif ext == '.ofd':
                    converted_pdf_path = read_ofd(str(file.resolve()))
                    loader = RapidOCRPDFLoader(file_path=converted_pdf_path)
                else:
                    loader = None

                if loader:
                    loaded = loader.load()
                    if converted_pdf_path:
                        os.remove(converted_pdf_path)
                    documents.extend([{"content":loaded_doc.page_content,"filename":file.stem,"filepath":str(file)} for loaded_doc in loaded])
            except (OSError, UnicodeDecodeError) as e:
                logger.error(f"Error reading file: {file}", e)
                logger.exception(e)
            except Exception as e:
                logger.error(f"Unhandled exception: {file}", e)
                logger.exception(e)
    return jsonify({"code":200,"data":documents})

def load(loader, texts):
    docs = loader.load()
    if docs:
        texts.extend([doc.page_content for doc in docs])
    return texts


if __name__ == '__main__':
    app.run()