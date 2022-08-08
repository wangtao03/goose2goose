import os
import time

import win32com.client


def pptx_to_mp4(pptx_path, mp4_target, resolution=1080, timeout=120):
    __status = 0
    if pptx_path == '' or mp4_target == '':
        return __status
    __start_time = time.time()
    __sdir = mp4_target[:mp4_target.rfind('\\')]
    if not os.path.exists(__sdir):
        os.makedirs(__sdir)

    __pptx = win32com.client.Dispatch('PowerPoint.Application')
    presentation = __pptx.Presentations.Open(pptx_path, WithWindow=True)
    presentation.CreateVideo(mp4_target, -1, 1, resolution)
    while True:
        try:
            time.sleep(0.1)
            if time.time() - __start_time > timeout:
                os.system("taskkill /f /im POWERPNT.EXE")
                __status = -1
                break
            if os.path.exists(mp4_path) and os.path.getsize(mp4_target) == 0:
                continue
            __status = 1
            break
        except FileNotFoundError:
            continue
        except Exception as __e:
            print('错误代码: {c}, 消息, {m}'.format(c=type(__e).__name__, m=str(__e)))
            break
    print(time.time() - __start_time)
    if __status != -1:
        __pptx.Quit()
    return __status


if __name__ == '__main__':

    for pptx in os.listdir("./videos"):
        if pptx.lower().endswith("pptx"):
            pptx_path = os.path.abspath(f"./videos/{pptx}")
            mp4_path = os.path.abspath(f"./videos/{pptx}.mp4")
            print(f"正在转换: {pptx}")
            status = 0
            try:
                status = pptx_to_mp4(pptx_path, mp4_path, 1080, 120)
            except Exception as e:
                print('错误代码: {c}, 消息, {m}'.format(c=type(e).__name__, m=str(e)))

            if status == -1:
                print('失败:转换超时!')
            elif status == 1:
                print('成功!')
            else:
                if os.path.exists(mp4_path):
                    os.remove(mp4_path)
                print('失败:未知格式的文件，请手动转换！')
