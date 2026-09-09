"""
未加工零件自动归档系统主入口

正式执行流程：
1. 加载配置
2. 连接 Google Drive
3. 扫描正在加工和拆图结果
4. 读取完成文件
5. 数据校验
6. 更新累计加工台账
7. 更新当前待加工零件
8. 归档已录入文件
9. 输出日志
"""

from modules.logger import get_logger
from modules.drive_manager import DriveManager

logger = get_logger()


def run():
    logger.info("开始执行未加工零件自动更新任务")

    drive = DriveManager()

    if not drive.check_config():
        raise RuntimeError("缺少 Google Drive 配置，请设置 GitHub Secrets")

    logger.info("Google Drive配置检查通过")

    # 后续生产模块调用：
    # 1. 读取Drive文件
    # 2. Excel解析
    # 3. 数据校验
    # 4. 累计加工计算
    # 5. 生成结果文件
    # 6. 上传并归档

    logger.info("云端执行环境已准备，等待授权后运行生产数据")


if __name__ == "__main__":
    run()
