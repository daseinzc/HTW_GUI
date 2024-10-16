import logging
from time import strftime
import combinemodule
import single_biaomodule
import excel_catchmodule

# 设置日志配置
logging.basicConfig(filename="main_log.txt", level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")

logging.info("开始读取Excel文件")
ls_hang, ls_school = excel_catchmodule.excelcatch()

if not ls_hang:
    logging.error("未从Excel文件中读取到任何数据")
    exit()

logging.info("成功读取Excel文件，开始生成单个文档")

# 逐个生成单个文档
for i in range(len(ls_hang)):
    try:
        school, year, month, money, due_day = ls_hang[i]
        formatted_date = due_day.strftime("%Y-%m-%d")
        ls_str = formatted_date.split('-')
        formatted_date = ls_str[0] + '年' + ls_str[1] + '月' + ls_str[2] + '日'
        single_biaomodule.single_conduct(str(school), str(year), str(month), str(money), str(formatted_date), i + 1)
        logging.info(f"成功生成文档：{i + 1} - {school}")
        print(f"单个文档生成进度：{i + 1}/{len(ls_hang)}")
    except Exception as e:
        logging.error(f"生成文档时出错：{e}")

logging.info("开始合并文档")
print("开始合并文档...")

# 调用合并模块
try:
    combinemodule.combine_docx()
    logging.info("文档合并完成")
    print("文档合并完成")
except Exception as e:
    logging.error(f"合并文档时出错：{e}")
    print("合并文档时出错，请查看日志")

input("按任意键退出...")
