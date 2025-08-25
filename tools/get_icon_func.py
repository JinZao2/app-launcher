import os
import sys
from PIL import Image
import win32api
import win32con
import win32gui
import win32ui


def extract_icon(exe_path):
    """
    提取指定程序的图标并返回PIL Image对象（更大尺寸并保留透明通道）
    """
    if not os.path.exists(exe_path):
        raise FileNotFoundError(f"文件不存在: {exe_path}")

    try:
        # 获取图标，增加图标尺寸到64x64
        large_icon, small_icon = win32gui.ExtractIconEx(exe_path, 0)

        if large_icon:
            hicon = large_icon[0]
        elif small_icon:
            hicon = small_icon[0]
        else:
            raise Exception("无法提取图标")

        # 绘制图标到设备上下文，使用64x64尺寸
        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()

        # 修改为64x64大小的图标
        icon_size = 64
        hbmp.CreateCompatibleBitmap(hdc, icon_size, icon_size)
        hdc = hdc.CreateCompatibleDC()

        old_bmp = hdc.SelectObject(hbmp)
        # 绘制图标
        win32gui.DrawIconEx(hdc.GetSafeHdc(), 0, 0, hicon, icon_size, icon_size, 0, None, win32con.DI_NORMAL)

        # 获取位图信息
        bmp_info = hbmp.GetInfo()
        bmp_str = hbmp.GetBitmapBits(True)

        # 使用RGBA模式保留透明通道
        img = Image.frombuffer(
            'RGBA',  # 改为RGBA模式保留透明度
            (bmp_info['bmWidth'], bmp_info['bmHeight']),
            bmp_str, 'raw', 'BGRA', 0, 1  # 对应BGRA格式
        )

        # 清理资源
        hdc.SelectObject(old_bmp)
        hdc.DeleteDC()
        win32gui.DestroyIcon(hicon)

        return img

    except Exception as e:
        print(f"提取图标时出错: {str(e)}")
        return None


if __name__ == "__main__":
    if len(sys.argv) > 1:
        exe_path = sys.argv[1]
    else:
        exe_path = input("请输入程序路径: ").replace("\"", "").strip()

    icon_image = extract_icon(exe_path)

    if icon_image:
        print("成功提取图标")
    else:
        print("提取图标失败")
