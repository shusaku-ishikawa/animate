import pipelines as pl
import openpyxl

if __name__ == "__main__":
    try:
        book = openpyxl.load_workbook('animate.xlsx', data_only = True)
    except Exception as e:
        raise Exception('エクセルファイルが開ませんでした')
    else:
        target_book = openpyxl.Workbook()
        queries = pl.QueryManager(book['Sheet1']).queries
        for q in queries:
            print(q)
       