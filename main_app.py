# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Nuclei扫描报告生成器 - 主程序入口（精简版）
功能：仅保留 Word 报告生成器，其他功能已移除
"""

import tkinter as tk
from tkinter import messagebox


class MainApp:
    """主程序首页（仅保留 Word 报告生成器）"""

    def __init__(self, root):
        self.root = root
        self.root.title("Nuclei扫描报告生成器")

        # 不再全屏，改为合适大小窗口（可手动最大化）
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)
        self.root.state('zoomed')  # Windows下最大化；macOS/Linux用 'zoomed' 或删除这行

        # 设置图标
        try:
            self.root.iconbitmap("nuclei.ico")
        except:
            pass

        self.setup_home_page()

    def setup_home_page(self):
        """设置精简首页界面"""
        # 清空
        for widget in self.root.winfo_children():
            widget.destroy()

        # 绑定 ESC 退出全屏（如果用户手动最大化了）
        self.root.bind('<Escape>', lambda e: self.root.destroy())

        # ==================== 顶部标题栏 ====================
        top_frame = tk.Frame(self.root, bg="#D9D9D9", height=80)
        top_frame.pack(fill="x")
        top_frame.pack_propagate(False)

        tk.Label(top_frame, text="Nuclei扫描报告生成器", font=("微软雅黑", 24, "bold"),
                 fg="#2c3e50", bg="#D9D9D9").pack(pady=20)

        # ==================== 主内容区 ====================
        content_frame = tk.Frame(self.root, bg="#f8f9fa")
        content_frame.pack(fill="both", expand=True, padx=50, pady=30)

        desc = tk.Label(content_frame,
                        text="• 支持IP地址和设备名称两种报告类型\n"
                             "• 批量处理多个文件，一键生成标准Word报告\n"
                             "•  自动填充检测时间、版本号、表格、样式\n"
                             "• 输出到项目根目录，文件名自动生成",
                        font=("微软雅黑", 14), fg="#34495e", bg="#f8f9fa", justify="left")
        desc.pack(pady=20)

        # 大按钮
        enter_btn = tk.Button(content_frame, text="进入 Word 报告生成器",
                              font=("微软雅黑", 18, "bold"),
                              bg="#9b59b6", fg="white",
                              width=30, height=3,
                              cursor="hand2",
                              command=self.open_word_report)
        enter_btn.pack(pady=60)

        # ==================== 底部信息 ====================
        footer = tk.Frame(self.root, bg="#ecf0f1", height=60)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)

        tk.Label(footer, text="版本 3.0（增强版）", font=("微软雅黑", 10), fg="#7f8c8d", bg="#ecf0f1") \
            .pack(side="left", padx=30, pady=15)
        tk.Label(footer, text="按 ESC 键退出程序", font=("微软雅黑", 10), fg="#7f8c8d", bg="#ecf0f1") \
            .pack(side="right", padx=30, pady=15)

    def open_word_report(self):
        """打开 Word 报告生成器"""
        try:
            from word_report_generator import WordReportGenerator
            self.root.withdraw()  # 隐藏主窗口
            word_window = tk.Toplevel(self.root)
            word_window.protocol("WM_DELETE_WINDOW", self.return_to_home)  # 关闭时返回首页
            WordReportGenerator(word_window, self)
        except ImportError as e:
            messagebox.showerror("错误", f"无法启动 Word 报告生成器：\n{e}")

    def return_to_home(self):
        """返回首页（当子窗口关闭时调用）"""
        self.root.deiconify()

    def toggle_fullscreen(self, event=None):
        """F11 切换全屏（可选）"""
        state = not self.root.attributes('-fullscreen')
        self.root.attributes('-fullscreen', state)


def main():
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
