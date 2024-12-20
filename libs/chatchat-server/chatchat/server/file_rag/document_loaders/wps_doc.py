from chatchat.server.file_rag.document_loaders.file_loader_service import read_wps

if __name__ == '__main__':
    r = read_wps(r"杭州英普环境技术股份有限公司(编号2477).wps")
    print(r)
    r = read_wps(r"公文智能排版与校对系统-管理员手册.doc")
    print(r)