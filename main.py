import time
import datetime
import os
from threading import Timer
import psutil

import autoex


def kill_process(name):
    # 获取当前所有应用的pids
    pids = psutil.pids()
    for pid in pids:
        # 根据pid 获取进程对象
        p = psutil.Process(pid)
        # 获取每个pid的应用对应的文件名字
        process_name = p.name()
        # print(process_name)
        # 判断形参是否是应用名字的子串来判断进程是否存在
        if name in process_name:
            # print("Process name is: %s, pid is: %s" % (process_name, pid))  # 1,33664
            try:
                # 如果存在，就删掉
                import subprocess
                subprocess.Popen("cmd.exe /k taskkill /F /T /PID %i" % pid, shell=True)
            except OSError:
                print('没有此进程!!!')


# 创建每隔5秒钟对应用的状态进行检测的定时器线程
# name是应用的名字
def restart_process(name):
    formattime = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    pids = psutil.pids()
    name_list = []
    for pid in pids:
        p = psutil.Process(pid)
        process_name = p.name()
        name_list.append(process_name)
    # 如果当前进程中已经没有了此应用
    if name not in name_list:
        # 声明变量的全局性
        global count
        count += 1
        print(formattime + "第" + str(count) + "次发现异常重连")
        os.system("start C:\\\"Program Files\"\\\"CAD Exchanger\"\\bin\\Exchanger.exe")
    global timer
    timer = Timer(10, restart_process, ("Exchanger.exe",))
    timer.start()


# 23个导出失败：程序崩溃
# 32号导入失败：点到滑动块的问题(已解决)
# 52号导出失败：程序崩溃
# 59导入失败：59号在逻辑上存在于第一页(已解决)
# 第115个文件导出失败：程序崩溃
# 第171个文件导出失败：展示到界面所需要的时间很长(已经解决，需要添加display是否成功的判断)
# 第199个文件导出失败：程序崩溃
# 第201个文件导入失败：无法导入
# 第204个文件导出失败：程序崩溃（已经解决程序崩溃的时候正常退出程序）
# 第321个文件在display的时候卡住，程序不响应（已经解决如果程序无响应就关闭程序程序）
# 322个文件在display的时候卡住，程序不响应
# 349个文件在display的时候卡住，程序不响应
# 462个文件导入失败
# 473个文件能导入，导出的时候发生崩溃

if __name__ == "__main__":
    # 创建并初始化计时器线程
    count = 0
    timer = Timer(10, restart_process, ("Exchanger.exe",))
    timer.start()

    for i in range(199, 204):
        try:
            # result = autoex.autoifc(i)
            result = autoex.autoobj(i)
            print("*******************************************")
            if result == -1:
                kill_process('Exchanger')
                time.sleep(2)
                continue
            # autoex.autoifc()

        except Exception as e:
            if repr(e).split("\'")[1] == "timed out":
                print("程序崩溃")
                kill_process('Exchanger')
            else:
                print(repr(e))
    timer.cancel()



