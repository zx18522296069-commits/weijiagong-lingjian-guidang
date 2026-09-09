"""
未加工零件自动归档系统
主入口

后续接入：
- Google Drive读取
- 拆图结果识别
- 累计加工计算
- Excel生成
- 已录入归档
"""

from datetime import datetime


def main():
    print(f"未加工零件任务启动: {datetime.now()}")
    # TODO: 接入正式业务逻辑


if __name__ == "__main__":
    main()
