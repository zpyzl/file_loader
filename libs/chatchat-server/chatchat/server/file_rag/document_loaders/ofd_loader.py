import base64
import os

from easyofd import OFD


def read_ofd(file_path):
    file_prefix = os.path.splitext(os.path.split(file_path)[1])[0]
    with open(file_path, "rb") as f:
        ofd_b64 = str(base64.b64encode(f.read()), "utf-8")
    ofd = OFD()  # 初始化OFD 工具类
    ofd.read(ofd_b64, save_xml=True, xml_name=f"{file_prefix}_xml")  # 读取ofdb64
    # print("ofd.data", ofd.data) # ofd.data 为程序解析结果
    pdf_bytes = ofd.to_pdf()  # 转pdf
    ofd.del_data()
    dir_path = os.path.split(file_path)[0]

    pdf_file_path = os.path.join(dir_path, f"{file_prefix}.pdf")
    with open(pdf_file_path, "wb") as f:
        f.write(pdf_bytes)
    return pdf_file_path

if __name__ == '__main__':
    file_path1 = r"D:\test_rag_doc\t\a.ofd"
    read_ofd(file_path1)