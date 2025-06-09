import os
import sys
import time
from pathlib import Path

from flask import Flask, request, jsonify
from tika import parser
from tqdm import tqdm

sys.path.append(os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))))

from chatchat.server.file_rag.document_loaders import RapidOCRDocLoader, RapidOCRPDFLoader
from chatchat.server.file_rag.document_loaders.ofd_loader import read_ofd
from log_utils import setup_log

logger = setup_log('file_loader_service.log',True)

app = Flask(__name__)

DIRPATH = r"D:\test_rag_doc"


@app.route('/load_file', methods=['GET', 'POST'])
def load_file():
    filepath = request.json.get('file_path')
    documents = []
    file = Path(filepath)
    start = time.time()
    code = 500
    try:
        logger.info(f"Processing {filepath}")
        ext = file.suffix
        print(f"{file.stat().st_size}")
        if ext in ['.docx']:
            loader = RapidOCRDocLoader(file_path=file,strategy='fast')
        elif ext == '.pdf':
            loader = RapidOCRPDFLoader(file_path=file,strategy='fast')
        # elif ext == '.ofd':
        #     converted_pdf_path = read_ofd(str(file.resolve()))
        #     loader = RapidOCRPDFLoader(file_path=converted_pdf_path)
        else:
            loader = None

        if loader:
            loaded = loader.load()
            # if converted_pdf_path:
            #     os.remove(converted_pdf_path)
            #     converted_pdf_path = None
            documents.extend([{"content":loaded_doc.page_content,"filename":file.stem,"filepath":str(file)} for loaded_doc in loaded])
            code = 200
        elif ext in ['.wps','.doc']:
            content = read_wps(filepath)
            documents.append({"content": content, "filename": file.stem, "filepath": str(file)})
            code = 200
        else:
            code = 500
    except (OSError, UnicodeDecodeError) as e:
        logger.error(f"Error reading file: {file}", e)
        logger.exception(e)
        code = 500
    except Exception as e:
        logger.error(f"Unhandled exception: {file}", e)
        logger.exception(e)
        code = 500
    logger.info(f"Load {filepath} cost {time.time() - start:.2f} seconds")
    return jsonify({"code":code,"data":documents})

def load(loader, texts):
    docs = loader.load()
    if docs:
        texts.extend([doc.page_content for doc in docs])
    return texts

def read_wps(file_path):
    parsed = parser.from_file(file_path)
    text = parsed.get('content','')
    return text


if __name__ == '__main__':
    app.run()