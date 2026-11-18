"""CustomTkinter 独立打包入口，复用主程序的桌面启动和 JSON 规则重放。"""

from main import main


if __name__ == "__main__":
    raise SystemExit(main(default_ui="customtkinter"))
