##这个版本，优化了图标显示和布局，支持缩放功能，添加了配置保存功能
"""
新功能说明
这个修改后的启动器具有以下特点：
- 图标更大且保持透明效果
- 取消水平滚动，图标自动换行
- 分类管理应用程序
- 支持搜索功能
- 右键菜单操作
- 支持缩放功能（Ctrl+滚轮或Ctrl++/-/0）
- 退出时保存缩放值、窗口大小和位置，下次启动时恢复

使用方法：
1. 运行程序后，会自动创建apps文件夹和几个示例分类
2. 在相应分类文件夹中放入应用程序的快捷方式（.lnk文件）
3. 按F5刷新启动器查看添加的应用
4. 使用Ctrl+鼠标滚轮或Ctrl++/-/0进行缩放
"""
import sys
import os
import subprocess
import pythoncom
import win32com.client
import json  # 新增：导入json模块处理配置文件
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QGridLayout,
                             QToolButton, QScrollArea, QVBoxLayout, QLineEdit,
                             QMenu, QInputDialog, QLabel, QFrame)
from PyQt6.QtGui import QIcon, QPixmap, QFont, QColor, QAction, QImage, QPainter, QWheelEvent
from PyQt6.QtCore import Qt, QSize, QPoint  # 新增：导入QPoint处理窗口位置

# 导入外部图标提取函数
from tools.get_icon_func import extract_icon


class AppLauncher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.app_categories = {}  # 按分类存储应用 {分类名: [应用列表]}
        self.favorite_apps = []
        self.apps_root_dir = "apps"  # 相对目录，存放分类文件夹
        self.scale_factor = 1.0  # 缩放因子，默认1.0
        self.min_scale = 0.5  # 最小缩放
        self.max_scale = 2.0  # 最大缩放
        self.config_path = os.path.join("profile", "config.json")  # 配置文件路径
        self.init_ui()

    def init_ui(self):
        # 设置窗口标题和默认大小
        self.setWindowTitle("JinZao的应用启动器")

        # 加载配置
        self.load_config()

        # 创建主部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # 添加搜索框
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("搜索应用...")
        self.search_box.textChanged.connect(self.filter_apps)
        self.search_box.setStyleSheet("""
            QLineEdit {
                border: 1px solid #CCCCCC;
                border-radius: 8px;
                padding: 8px 12px;
                font-size: 14px;
                background-color: rgba(255, 255, 255, 0.8);
            }
        """)
        main_layout.addWidget(self.search_box)

        # 创建滚动区域用于显示应用图标（只垂直滚动）
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        # 禁用水平滚动条
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: rgba(249, 250, 251, 0.9);
            }
        """)

        # 滚动区域内的部件 - 使用垂直布局来容纳多个分类
        self.scroll_content = QWidget()
        self.main_content_layout = QVBoxLayout(self.scroll_content)
        self.main_content_layout.setSpacing(20)
        # 调整主内容边距，增加右边距避免被滚动条遮挡
        self.main_content_layout.setContentsMargins(10, 10, 25, 10)

        self.scroll_area.setWidget(self.scroll_content)
        main_layout.addWidget(self.scroll_area)

        # 确保应用目录和配置目录存在
        self.ensure_apps_directory_exists()
        self.ensure_profile_directory_exists()

        # 加载应用程序
        self.load_applications()

        # 设置窗口样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F9FAFB;
            }
        """)

        # 显示窗口
        self.show()

    def ensure_profile_directory_exists(self):
        """确保存放配置文件的目录存在"""
        profile_dir = os.path.dirname(self.config_path)
        if not os.path.exists(profile_dir):
            os.makedirs(profile_dir)

    def ensure_apps_directory_exists(self):
        """确保存放应用分类的目录存在"""
        if not os.path.exists(self.apps_root_dir):
            os.makedirs(self.apps_root_dir)
            # 创建几个示例分类文件夹
            example_categories = ["办公软件", "浏览器", "开发工具", "娱乐"]
            for category in example_categories:
                os.makedirs(os.path.join(self.apps_root_dir, category), exist_ok=True)

    def load_config(self):
        """加载保存的配置"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                    # 恢复窗口位置
                    if "pos" in config:
                        self.move(QPoint(config["pos"][0], config["pos"][1]))

                    # 恢复窗口大小
                    if "size" in config:
                        self.resize(config["size"][0], config["size"][1])
                    else:
                        self.setGeometry(100, 100, 1000, 700)  # 默认大小

                    # 恢复缩放因子
                    if "scale_factor" in config:
                        # 确保缩放因子在有效范围内
                        self.scale_factor = max(self.min_scale,
                                                min(self.max_scale, config["scale_factor"]))
        except Exception as e:
            print(f"加载配置失败: {e}")
            # 使用默认设置
            self.setGeometry(100, 100, 1000, 700)

    def save_config(self):
        """保存当前配置"""
        try:
            # 确保配置目录存在
            self.ensure_profile_directory_exists()

            config = {
                "pos": [self.x(), self.y()],  # 窗口位置
                "size": [self.width(), self.height()],  # 窗口大小
                "scale_factor": self.scale_factor  # 缩放因子
            }

            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")

    def closeEvent(self, event):
        """窗口关闭时保存配置"""
        self.save_config()
        event.accept()

    def load_applications(self):
        """从相对目录加载分类好的应用程序"""
        self.app_categories.clear()

        # 遍历应用根目录下的所有文件夹（分类）
        for category in os.listdir(self.apps_root_dir):
            category_path = os.path.join(self.apps_root_dir, category)

            # 只处理目录
            if os.path.isdir(category_path):
                apps_in_category = []

                # 遍历分类目录下的所有快捷方式
                for file in os.listdir(category_path):
                    file_path = os.path.join(category_path, file)

                    # 只处理.lnk快捷方式文件
                    if file.lower().endswith('.lnk'):
                        try:
                            # 解析快捷方式
                            lnk_info = self.parse_shortcut(file_path)
                            if lnk_info:
                                app_name = os.path.splitext(file)[0]  # 从文件名获取应用名
                                target_path = lnk_info['target']

                                # 获取图标
                                icon = self.get_app_icon(target_path)

                                apps_in_category.append({
                                    'name': app_name,
                                    'path': file_path,  # 快捷方式路径
                                    'target': target_path,  # 实际应用路径
                                    'icon': icon,
                                    'category': category
                                })
                        except Exception as e:
                            print(f"处理快捷方式 {file} 时出错: {e}")

                if apps_in_category:  # 只添加有应用的分类
                    self.app_categories[category] = apps_in_category

        # 加载收藏的应用
        self.load_favorite_apps()

        # 显示应用图标
        self.display_apps()

    def parse_shortcut(self, lnk_path):
        """解析Windows快捷方式(.lnk)文件"""
        try:
            pythoncom.CoInitialize()
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(lnk_path)
            return {
                'target': shortcut.Targetpath,
                'arguments': shortcut.Arguments,
                'working_dir': shortcut.WorkingDirectory
            }
        except Exception as e:
            print(f"解析快捷方式失败: {e}")
            return None
        finally:
            pythoncom.CoUninitialize()

    def get_app_icon(self, exe_path):
        """从可执行文件获取图标，调用外部extract_icon函数"""
        try:
            if os.path.exists(exe_path) and exe_path.lower().endswith('.exe'):
                # 调用外部函数获取PIL Image对象
                pil_image = extract_icon(exe_path)
                if pil_image:
                    # 保持alpha通道（透明）
                    img = pil_image.convert("RGBA")
                    data = img.tobytes("raw", "RGBA")
                    q_image = QImage(data, img.width, img.height, QImage.Format.Format_RGBA8888)

                    # 转换为QPixmap并创建QIcon
                    pixmap = QPixmap.fromImage(q_image)
                    return QIcon(pixmap)
        except Exception as e:
            print(f"获取图标失败: {e}")

        # 如果无法获取图标，返回带透明背景的默认图标
        default_pixmap = QPixmap(int(64 * self.scale_factor), int(64 * self.scale_factor))
        default_pixmap.fill(Qt.GlobalColor.transparent)  # 设置透明背景
        painter = QPainter(default_pixmap)
        painter.setPen(QColor(200, 200, 200))
        painter.drawRect(10, 10, int(44 * self.scale_factor), int(44 * self.scale_factor))  # 绘制简单边框作为默认图标
        painter.end()
        return QIcon(default_pixmap)

    def display_apps(self, filtered_apps=None):
        """显示应用程序图标，按分类组织"""
        # 清除现有内容
        while self.main_content_layout.count():
            item = self.main_content_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        if filtered_apps:
            # 处理搜索过滤的情况，按分类重新组织
            filtered_categories = {}
            for app in filtered_apps:
                if app['category'] not in filtered_categories:
                    filtered_categories[app['category']] = []
                filtered_categories[app['category']].append(app)
            categories_to_display = filtered_categories
        else:
            # 显示所有分类
            categories_order = sorted(self.app_categories.keys())  # 按名称排序分类
            categories_to_display = {cat: self.app_categories[cat] for cat in categories_order}

        # 按分类显示应用
        for category, apps in categories_to_display.items():
            # 添加分类标题
            category_label = QLabel(category)
            category_label.setFont(QFont("SimHei", int(14 * self.scale_factor), QFont.Weight.Bold))
            category_label.setStyleSheet("color: #333333; margin-top: 10px;")
            self.main_content_layout.addWidget(category_label)

            # 添加分隔线
            line = QFrame()
            line.setFrameShape(QFrame.Shape.HLine)
            line.setFrameShadow(QFrame.Shadow.Sunken)
            line.setStyleSheet("background-color: #E5E7EB; height: 1px;")
            self.main_content_layout.addWidget(line)

            # 创建网格布局显示该分类下的应用
            category_widget = QWidget()
            grid_layout = QGridLayout(category_widget)
            grid_layout.setSpacing(int(15 * self.scale_factor))
            # 增加右 margin 避免被滚动条遮挡
            grid_layout.setContentsMargins(5, 5, 15, 15)

            # 根据缩放因子和窗口宽度计算列数
            base_item_width = 160  # 基础宽度
            scaled_item_width = base_item_width * self.scale_factor
            cols = max(1, int(self.width() // scaled_item_width))

            # 添加应用按钮
            for i, app in enumerate(apps):
                row = i // cols
                col = i % cols

                # 创建应用按钮
                app_button = QToolButton()
                app_button.setIcon(app['icon'])
                # 根据缩放因子调整图标大小
                icon_size = int(64 * self.scale_factor)
                app_button.setIconSize(QSize(icon_size, icon_size))
                app_button.setText(app['name'])
                app_button.setToolTip(app['target'])
                app_button.setObjectName(app['path'])

                # 绑定点击事件（启动应用）
                app_button.clicked.connect(
                    lambda checked, target=app['target'], args=app.get('arguments', ''):
                    self.launch_app(target, args)
                )

                # 计算按钮尺寸
                btn_min_width = int(120 * self.scale_factor)
                btn_max_width = int(140 * self.scale_factor)
                btn_height = int(140 * self.scale_factor)

                # 设置按钮样式 - 根据缩放因子调整
                app_button.setStyleSheet(f"""
                    QToolButton {{
                        background-color: rgba(255, 255, 255, 0.7);
                        border: 1px solid transparent;
                        border-radius: {int(8 * self.scale_factor)}px;
                        padding: {int(8 * self.scale_factor)}px;
                        margin: {int(5 * self.scale_factor)}px;
                        text-align: center;
                        font-size: {int(12 * self.scale_factor)}px;
                        min-width: {btn_min_width}px;
                        max-width: {btn_max_width}px;
                        min-height: {btn_height}px;
                        max-height: {btn_height}px;
                        white-space: normal;
                        word-wrap: break-word;
                    }}
                    QToolButton:hover {{
                        background-color: rgba(240, 240, 240, 0.9);
                        border: 1px solid #E5E7EB;
                    }}
                    QToolButton:pressed {{
                        background-color: rgba(220, 220, 220, 0.9);
                    }}
                """)

                # 设置文本在图标下方
                app_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)

                # 添加右键菜单
                app_button.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
                app_button.customContextMenuRequested.connect(
                    lambda pos, app=app, widget=app_button:
                    self.show_context_menu(pos, app, widget)
                )

                grid_layout.addWidget(app_button, row, col)

            self.main_content_layout.addWidget(category_widget)

        # 添加一个伸缩项，将所有内容推到顶部
        self.main_content_layout.addStretch()

    def launch_app(self, target_path, arguments=""):
        """启动应用程序"""
        try:
            if arguments:
                subprocess.Popen([target_path, arguments])
            else:
                subprocess.Popen([target_path])
        except Exception as e:
            print(f"无法启动应用: {e}")

    def filter_apps(self, text):
        """根据搜索文本过滤应用"""
        if not text or text.strip() == "":
            self.display_apps()
            return

        # 搜索所有分类中的应用
        all_apps = []
        for apps in self.app_categories.values():
            all_apps.extend(apps)

        # 过滤应用
        filtered = [app for app in all_apps if text.lower() in app['name'].lower()]
        self.display_apps(filtered)

    def show_context_menu(self, position, app, widget):
        """显示右键菜单"""
        menu = QMenu()

        # 启动应用动作
        launch_action = QAction("启动", self)
        launch_action.triggered.connect(
            lambda: self.launch_app(app['target'], app.get('arguments', ''))
        )

        # 重命名动作
        rename_action = QAction("重命名", self)
        rename_action.triggered.connect(
            lambda: self.rename_app(app, widget)
        )

        # 添加到收藏动作
        favorite_action = QAction(
            "添加到收藏" if app['path'] not in self.favorite_apps else "从收藏移除",
            self
        )
        favorite_action.triggered.connect(
            lambda: self.toggle_favorite(app['path'])
        )

        # 删除动作
        remove_action = QAction("从启动器移除", self)
        remove_action.triggered.connect(
            lambda: self.remove_app(app)
        )

        # 添加到菜单
        menu.addAction(launch_action)
        menu.addAction(rename_action)
        menu.addAction(favorite_action)
        menu.addSeparator()
        menu.addAction(remove_action)

        # 在鼠标位置显示菜单
        menu.exec(widget.mapToGlobal(position))

    def rename_app(self, app, widget):
        """重命名应用（实际是重命名快捷方式文件）"""
        new_name, ok = QInputDialog.getText(
            self, "重命名应用", "新名称:", text=app['name']
        )

        if ok and new_name and new_name != app['name']:
            try:
                # 获取文件扩展名
                ext = os.path.splitext(app['path'])[1]
                # 构建新路径
                new_path = os.path.join(os.path.dirname(app['path']), new_name + ext)

                # 重命名文件
                os.rename(app['path'], new_path)

                # 更新应用信息
                app['name'] = new_name
                app['path'] = new_path
                widget.setText(new_name)

                # 重新加载应用列表
                self.load_applications()
            except Exception as e:
                print(f"重命名失败: {e}")

    def toggle_favorite(self, app_path):
        """添加或移除收藏"""
        if app_path in self.favorite_apps:
            self.favorite_apps.remove(app_path)
        else:
            self.favorite_apps.append(app_path)
        self.save_favorite_apps()

    def remove_app(self, app):
        """从启动器移除应用（删除快捷方式）"""
        try:
            if os.path.exists(app['path']):
                os.remove(app['path'])

            # 重新加载应用
            self.load_applications()
        except Exception as e:
            print(f"删除应用失败: {e}")

    def load_favorite_apps(self):
        """从文件加载收藏的应用"""
        try:
            if os.path.exists('favorites.txt'):
                with open('favorites.txt', 'r', encoding='utf-8') as f:
                    self.favorite_apps = [line.strip() for line in f.readlines()]
        except Exception as e:
            print(f"加载收藏失败: {e}")

    def save_favorite_apps(self):
        """保存收藏的应用到文件"""
        try:
            with open('favorites.txt', 'w', encoding='utf-8') as f:
                for app_path in self.favorite_apps:
                    f.write(f"{app_path}\n")
        except Exception as e:
            print(f"保存收藏失败: {e}")

    def resizeEvent(self, event):
        """窗口大小改变时重新排列图标"""
        self.display_apps()
        super().resizeEvent(event)

    def keyPressEvent(self, event):
        """处理键盘事件"""
        # 缩放控制
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            if event.key() == Qt.Key.Key_Plus or event.key() == Qt.Key.Key_Equal:
                self.scale_factor = min(self.max_scale, self.scale_factor + 0.1)
                self.display_apps()
                return
            elif event.key() == Qt.Key.Key_Minus:
                self.scale_factor = max(self.min_scale, self.scale_factor - 0.1)
                self.display_apps()
                return
            elif event.key() == Qt.Key.Key_0:
                self.scale_factor = 1.0
                self.display_apps()
                return

        if event.key() == Qt.Key.Key_Escape:
            self.close()
        elif event.key() == Qt.Key.Key_F5:
            self.load_applications()
        else:
            super().keyPressEvent(event)

    def wheelEvent(self, event: QWheelEvent):
        """处理鼠标滚轮事件实现缩放"""
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            if delta > 0:
                # 滚轮向上，放大
                self.scale_factor = min(self.max_scale, self.scale_factor + 0.1)
            else:
                # 滚轮向下，缩小
                self.scale_factor = max(self.min_scale, self.scale_factor - 0.1)
            self.display_apps()
            event.accept()
        else:
            super().wheelEvent(event)


if __name__ == '__main__':
    # 确保中文显示正常
    font = QFont("SimHei")

    app = QApplication(sys.argv)
    app.setFont(font)

    # 设置全局样式
    app.setStyle("Fusion")

    launcher = AppLauncher()
    sys.exit(app.exec())