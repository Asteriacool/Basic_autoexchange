### 一、准备

#### 1.Spy++

(https://zhuanlan.zhihu.com/p/355878952)

(1) 下载安装Spy++

(2) 使用Spy++进行窗体所属程序检测

* 使用到的Spy++的功能：或许Windows下任意窗口的句柄信息，进而获取到该窗口所属的应用程序
* 检测步骤
  * 在Spy菜单中选择Find Window
  * 拖动FinderTool到需要识别的窗口上

#### 2.Inspect

backend是windows上受支持的辅助功能技术

* Win 32 API(backend="win32")
* MS UI Automation(backend="uia")

对于程序适用于哪种backend，可以借助于GUI对象检查工具来进行检查

### 二、利用psutil库对于目前正在进行的进程进行判断

```python
import psutil

def proc_exist(process_name):
    p1 = psutil.pids() # 获取当前所有进程
    for pid in p1:
        # 判断搜索的进程是否在当前进程内
        if psutil.Process(pid).name() == process_name:
            return pid
```

```python
if isinstance(proc_exist('Exchanger.exe'), int):
    # 如果正在运行
else:
    # 如果没有运行
```

### 三、对应用程序进行操作

#### 1.连接到应用程序

使用pywinauto.application.Application()来创建一个应用程序对象

* start()用来在这个程序没有运行但是需要启动它的时候

```python
app = Application().start(r"c:\path\to\your\application -a -n -y --arguments")
```

* connect（）连接已经打开了的应用程序，指定下列选项之一
  * 应用进程的ID
    *  ` app = Application().connect(process=2341)`
  * 应用程序的窗口句柄
    * `app = Application().connect(handle=0x010f0c)`
  * 路径
    * `app = Application().connect （path = r “c：\ windows \ system32 \ notepad.exe” ）`

#### 2.窗口选择

首先选中窗口，才能对该窗口进行操作

```python
app.window(**kwargs)
```

* 返回的内容是一个WindowSpecification对象
* kwargs包含窗口属性的一系列筛选条件
  * class_name=None, # 类名
  * title=None, # 控件的标题文字，对应inspect中Name字段
  * control_type=None, # 控件类型，inspect界面LocalizedControlType字段的英文名
  * auto_id=None, # 这个也是固定的可以用，inspect界面AutomationId字段，但是很多控件没有这个属性

#### 3.窗口的操作

* window.descendants(**kwargs) 
  * 返回符合条件的所有后代元素列表,是BaseWrapper对象（或子类）
  * 用来判断在文件列表中是否存在某个文件，以及窗口中是否存在某个按钮元素（返回列表的非空性）

#### 4.控件的操作

* child_window(**kwargs)
  * 可以不管层级的找后代中某个符合条件的元素，最常用
  * 通过窗口选择控件
* ctrl.draw_outline(colour='green')
  * 控件外围画框，便于查看，支持'red', 'green', 'blue'
* ctrl.exists()
  * 对控件的存在性进行判断

### 四、对应用当前所处的不同状态进行判断和特殊情况的分析

### 【基础框架的实现】

![Image](C:\Users\aster\Desktop\pywinauto_ex\Image.png)

### 【自动化处理的实现】

需求：把目标文件夹中的所有文件都转换成IFC格式

总体思路通过id值依次获取窗口中的所有文件，如果当前窗口存在该文件，就直接点选；如果当前窗口不存在该文件，就通过鼠标滚动页面，并不断判断当前窗口中是否存在该文件，直到完成对最后一个文件的处理

细化问题与解决方案

* 文件路径的问题文件路径默认是上一次的打开路径（能不能选中指定路径？）
* 文件格式的问题（如果是文件夹该怎么办？）
* 如何找到所有的需要转换的文件
  * 如果只用descendants则只能得到当前窗口中的所有文件，因此考虑加入scroll函数

* 文件导出路径的问题文件路径默认是上一次的打开路径，需要自己新建文件夹（能不能选中指定路径？）
* 导入问题：某个id的文件由于出现在窗口的底部而被上层的滚动条覆盖，导致无法点选
  * 思路：在点击之前首先确定该文件的位置和底部水平滚动条的位置，看两者是否重合——如果重合，就加一次滚动，再点击，如果不重合就直接点击

```Python
bar_box = window.child_window(title="水平滚动条", auto_id="HorizontalScrollBar", control_type="ScrollBar")
rect = bar_box.rectangle()
rect_h = (rect.top + rect.bottom) / 2
# 目标文件
item = window.child_window(auto_id=str(id), control_type="ListItem")
while not item.exists():
    mouse_scroll(control=window, distance=-5)
    item = window.child_window(auto_id=str(id), control_type="ListItem")
rect_i = item.rectangle()
recti_h = (rect_i.top + rect_i.bottom) / 2


if abs(rect_h-recti_h) <= 3:
    mouse_scroll(control=window, distance=-5)
```

* 呈现问题：若成功，不同的文件渲染呈现所经历的时间不同；若失败，会提示渲染失败
  * 解决思路，利用wait函数，等待渲染完成，如果渲染不成功，则会报出超时的错误，在异常信息的处理位置处理

```python
# dlg.wait(wait_for, timeout=None, retry_interval=None)  # 等待窗口进入特定状态
# dlg.wait_not(wait_for_not, timeout=None, retry_interval=None)  # 等待窗口退出特定状态，即等待消失
# wait_for/wait_for_not:
#   * 'exists' means that the window is a valid handle
#   * 'visible' means that the window is not hidden
#   * 'enabled' means that the window is not disabled
#   * ready' means that the window is visible and enabled
#   * 'active' means that the window is active
# timeout: 设置超时的时间，如果在n秒后没有进入指定状态，则退出等待
# retry_interval:timeout内重试时间，How long to sleep between each retry.
# eg: dlg.wait('ready')
```

* 渲染问题：对于导入后在渲染阶段是的应用卡住不响应
  * 思路：利用psutil关闭应用进程

```python
def kill_process(name):
    pids = psutil.pids()
    for pid in pids:
        p = psutil.Process(pid)
        process_name = p.name()
        # print(process_name)
        if name in process_name:
            # print("Process name is: %s, pid is: %s" % (process_name, pid))  # 1,33664
            try:
                import subprocess
                subprocess.Popen("cmd.exe /k taskkill /F /T /PID %i" % pid, shell=True)
            except OSError:
                print('没有此进程!!!')
```

* 实现对于文件夹的文件进行自动重启
  * 对于可以预见的异常进行自动重启
    * 思路：对于导入导出失败的函数返回-1，如果是-1就终止程序，并重新启动
  * 应对程序中不可知原因的忽然中断
    * 思路：在主函数中加入计时器线程，每10秒钟检测目标应用程序是否存在
    * 10秒的解释：该应用程序的启动时间很长，避免在应用程序启动期间第二次检测不到又重开一个新的窗口，导致后期的误判定时器
    * 参考：[cnblogs.com/btc1996/p/14456264.html](http://cnblogs.com/btc1996/p/14456264.html)

```python
# 创建每隔5秒钟对应用的状态进行检测的定时器线程
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
        global count
        count += 1
        print(formattime + "第" + str(count) + "次发现异常重连")
        os.system("start C:\\\"Program Files\"\\\"CAD Exchanger\"\\bin\\Exchanger.exe")
        print("重启连接成功")
    global timer
    timer = Timer(5, restart_process, ("Exchanger.exe",))
    timer.start()
```

