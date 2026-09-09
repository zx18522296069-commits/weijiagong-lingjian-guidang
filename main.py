"""
未加工零件自动归档系统主入口

执行流程：
1. 加载配置
2. 连接 Google Drive
3. 扫描订单和完成文件
4. 读取 Excel
5. 数据校验
6. 累计加工计算
7. 生成结果文件
8. 上传结果
9. 归档已录入文件
10. 输出日志
"""

from modules.logger import get_logger
from modules.drive_manager import DriveManager
from modules.excel_reader import read_excel, find_completion_files
from modules.process_parts import calculate_remaining

logger = get_logger()


def run():
    logger.info("开始执行未加工零件自动更新任务")

    drive = DriveManager()

    if not drive.check_config():
        raise RuntimeError("缺少 Google Drive 配置，请设置 GitHub Secrets")

    logger.info("Google Drive配置检查通过")

    # 生产环境执行顺序
    # 1. 获取Drive文件列表
    # 2. 读取正在加工数据
    # 3. 读取拆图完成数据
    # 4. 按板材编号累计
    # 5. 生成当前待加工零件
    # 6. 更新累计加工台账
    # 7. 完成订单后归档

    logger.info("主流程框架已连接，等待真实Drive授权测试")


if __name__ == "__main__":
    run()
