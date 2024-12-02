import logging
from single_biaomodule import single_conduct
import excel_catchmodule
from combinemodule import combine_doc
import sys

# 设置日志配置
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('program.log'),  # 将日志写入文件
        logging.StreamHandler(sys.stdout)    # 将日志输出到控制台
    ]
)

def read_excel_data():
    """读取Excel数据并进行验证"""
    logging.info("开始读取Excel文件")
    try:
        ls_hang, ls_school = excel_catchmodule.excelcatch()

        # 验证是否读取到数据
        if not ls_hang:
            logging.error("未从Excel文件中读取到任何数据")
            sys.exit(1)  # 退出程序，状态码1表示错误

        logging.info(f"成功读取Excel文件，读取到 {len(ls_hang)} 行数据")
        return ls_hang

    except Exception as e:
        logging.error(f"读取Excel文件时发生错误：{e}")
        sys.exit(1)  # 退出程序，状态码1表示错误

def generate_single_documents(ls_hang):
    """逐行生成单个文档"""
    logging.info("开始生成单个文档")

    for i, row in enumerate(ls_hang):
        try:
            xuhao, school, money, month, year, due_day = row

            # 格式化日期和其他数据（如需要）

            # 调用单个文档生成函数
            single_conduct(str(school), str(year), str(month), str(money), str(due_day), i + 1)

            logging.info(f"成功生成文档：{i + 1} - {school}")
            print(f"单个文档生成进度：{i + 1}/{len(ls_hang)}")

        except Exception as e:
            logging.error(f"生成文档时出错：{e}, 文档编号: {i + 1}")

def main():
    """主程序入口"""
    try:
        # 读取 Excel 数据
        ls_hang = read_excel_data()

        # 生成单个文档
        generate_single_documents(ls_hang)

        # 合并文档
        logging.info("开始合并文档")
        combine_doc()

        logging.info("文档生成和合并完成")

    except Exception as e:
        logging.critical(f"程序执行出错：{e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
